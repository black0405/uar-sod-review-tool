/* =====================================================================
   SoD / UAR review — Oracle EBS R12 query stack
   All statements are SELECT-only. Run as a read-only user with SELECT
   on the APPS synonyms. Export each result to .xlsx with headers as-is.

   Evidence to keep per run (IPE): query text, parameters, run timestamp,
   instance name, row count, and a screenshot of the run.

   Q1-Q2 feed the tool (Q1 carries functions, which switches on Level 3).
   Q3-Q5 are validation queries for findings.
   ===================================================================== */


/* ---------------------------------------------------------------------
   Q1  Oracle access extract, with functions             -> tool slot 1
   One row per user x responsibility x function. The tool reads the
   responsibility columns for Level 1 SoD and the function columns for
   Level 3 (removing the duplicate responsibility rows itself).
   - FND_USER_RESP_GROUPS_ALL includes responsibilities inherited via
     roles (UMX), not only direct grants.
   - "End Date" is the effective end: earliest of assignment, user
     account and responsibility end dates; NULL = still open.
   - Functions: full menu tree per responsibility, after function (F)
     and menu (M) exclusions; read-only (QUERY_ONLY=YES) form variants
     are dropped. Responsibilities whose menus grant no
     function still appear, with blank function columns.
   - To run Level 1 only, drop the rfunc join and the two function columns.
   --------------------------------------------------------------------- */
WITH asg AS (
  SELECT DISTINCT
         fu.user_name, fu.email_address, furg.start_date,
         NULLIF(LEAST(NVL(furg.end_date, DATE '4712-12-31'),
                      NVL(fu.end_date,   DATE '4712-12-31'),
                      NVL(fr.end_date,   DATE '4712-12-31')),
                DATE '4712-12-31')  end_date,
         fr.responsibility_id, fr.application_id, fr.menu_id, fr.responsibility_key,
         frt.responsibility_name, fa.application_short_name, fu.last_logon_date
    FROM fnd_user_resp_groups_all  furg
    JOIN fnd_user                  fu  ON fu.user_id            = furg.user_id
    JOIN fnd_responsibility        fr  ON fr.responsibility_id  = furg.responsibility_id
                                      AND fr.application_id     = furg.responsibility_application_id
    JOIN fnd_responsibility_tl     frt ON frt.responsibility_id = fr.responsibility_id
                                      AND frt.application_id    = fr.application_id
                                      AND frt.language          = USERENV('LANG')
    JOIN fnd_application           fa  ON fa.application_id     = fr.application_id
),
tree AS (
  SELECT CONNECT_BY_ROOT me.menu_id                  root_menu_id,
         me.function_id, me.grant_flag,
         SYS_CONNECT_BY_PATH(me.menu_id, '/') || '/'  menu_path
    FROM fnd_menu_entries me
   START WITH me.menu_id IN (SELECT menu_id FROM asg)
 CONNECT BY NOCYCLE PRIOR me.sub_menu_id = me.menu_id
),
rfunc AS (  -- functions reachable per responsibility (computed once, not per user)
  SELECT DISTINCT r.responsibility_id, r.application_id,
         fff.function_name, ffft.user_function_name
    FROM (SELECT DISTINCT responsibility_id, application_id, menu_id FROM asg) r
    JOIN tree t                     ON t.root_menu_id   = r.menu_id
                                   AND t.grant_flag     = 'Y'
    JOIN fnd_form_functions    fff  ON fff.function_id  = t.function_id
    JOIN fnd_form_functions_tl ffft ON ffft.function_id = fff.function_id
                                   AND ffft.language    = USERENV('LANG')
   WHERE NOT REGEXP_LIKE(NVL(fff.parameters, ' '), 'QUERY_ONLY *= *"?YES', 'i')  -- read-only form variants
     AND NOT EXISTS (SELECT 1 FROM fnd_resp_functions rf
                      WHERE rf.responsibility_id = r.responsibility_id
                        AND rf.application_id    = r.application_id
                        AND (   (rf.rule_type = 'F' AND rf.action_id = fff.function_id)
                             OR (rf.rule_type = 'M' AND INSTR(t.menu_path, '/' || rf.action_id || '/') > 0)))
)
SELECT a.user_name               "User Name",
       a.email_address           "EMAIL_ADDRESS",
       a.responsibility_name     "Responsibility Name",
       a.start_date              "Start Date",
       a.end_date                "End Date",
       a.responsibility_key      "Responsibility Key",
       a.application_short_name  "Application Short Name",
       a.last_logon_date         "Last Logon",
       f.function_name           "Function Code",
       f.user_function_name      "Function Name"
  FROM asg a
  LEFT JOIN rfunc f ON f.responsibility_id = a.responsibility_id
                   AND f.application_id    = a.application_id
 ORDER BY 1, 3, 9;


/* ---------------------------------------------------------------------
   Q2  Person -> Oracle user mapping                     -> tool slot 2
   The tool links an Oracle user to a person when the username equals
   the local part of EMAIL_ADDRESS. This version builds that key from
   FND_USER.EMPLOYEE_ID, the link Oracle itself holds, so it works even
   when usernames do not follow the email naming convention.
   (Plain alternative: select papf.email_address as-is.)
   --------------------------------------------------------------------- */
SELECT papf.person_id                                   "PERSON_ID",
       papf.last_name                                   "LAST_NAME",
       CASE WHEN fu.user_name IS NOT NULL
            THEN LOWER(fu.user_name) || '@'
                 || NVL(SUBSTR(papf.email_address, INSTR(papf.email_address, '@') + 1), 'oracle.local')
       END                                              "EMAIL_ADDRESS",
       papf.full_name                                   "FULL_NAME"
  FROM per_all_people_f papf
  LEFT JOIN fnd_user    fu ON fu.employee_id = papf.person_id
 WHERE TRUNC(SYSDATE) BETWEEN papf.effective_start_date AND papf.effective_end_date
 ORDER BY papf.full_name;


/* ---------------------------------------------------------------------
   Q3  Function-level drill-down for one responsibility      (validation)
   Confirms a flagged responsibility really carries the conflicting
   functions. Walks the full menu tree and removes function and menu
   exclusions defined on the responsibility.
   --------------------------------------------------------------------- */
WITH resp AS (
  SELECT fr.responsibility_id, fr.application_id, fr.menu_id, frt.responsibility_name
    FROM fnd_responsibility    fr
    JOIN fnd_responsibility_tl frt ON frt.responsibility_id = fr.responsibility_id
                                  AND frt.application_id    = fr.application_id
                                  AND frt.language          = USERENV('LANG')
   WHERE frt.responsibility_name = :p_responsibility_name
),
excl AS (
  SELECT rf.rule_type, rf.action_id
    FROM fnd_resp_functions rf
    JOIN resp r ON r.responsibility_id = rf.responsibility_id
               AND r.application_id    = rf.application_id
),
menu_tree AS (
  SELECT me.function_id, me.grant_flag
    FROM fnd_menu_entries me
   START WITH me.menu_id = (SELECT menu_id FROM resp)
 CONNECT BY NOCYCLE PRIOR me.sub_menu_id = me.menu_id
         AND me.menu_id NOT IN (SELECT action_id FROM excl WHERE rule_type = 'M')
)
SELECT DISTINCT
       r.responsibility_name     "Responsibility",
       fff.function_name         "Function Code",
       ffft.user_function_name   "Function Name"
  FROM menu_tree mt
  JOIN fnd_form_functions    fff  ON fff.function_id  = mt.function_id
  JOIN fnd_form_functions_tl ffft ON ffft.function_id = fff.function_id
                                 AND ffft.language    = USERENV('LANG')
 CROSS JOIN resp r
 WHERE mt.grant_flag = 'Y'
   AND fff.function_id NOT IN (SELECT action_id FROM excl WHERE rule_type = 'F')
   AND NOT REGEXP_LIKE(NVL(fff.parameters, ' '), 'QUERY_ONLY *= *"?YES', 'i')  -- read-only form variants
 ORDER BY 3;


/* ---------------------------------------------------------------------
   Q4  Conflict exercised: same user entered the invoice AND the payment
       (AP_INVOICE_ENTRY x AP_PAYMENTS)                       (validation)
   For payment process requests, CREATED_BY on the check is the user who
   submitted the run.
   --------------------------------------------------------------------- */
SELECT fu.user_name           "Oracle User",
       asp.vendor_name        "Supplier",
       aia.invoice_num        "Invoice",
       aia.invoice_amount     "Invoice Amount",
       aca.check_number       "Payment Doc",
       aca.check_date         "Payment Date",
       aca.amount             "Payment Amount"
  FROM ap_invoices_all         aia
  JOIN ap_invoice_payments_all aipa ON aipa.invoice_id = aia.invoice_id
  JOIN ap_checks_all           aca  ON aca.check_id    = aipa.check_id
  JOIN ap_suppliers            asp  ON asp.vendor_id   = aia.vendor_id
  JOIN fnd_user                fu   ON fu.user_id      = aia.created_by
 WHERE aca.created_by = aia.created_by
   AND aca.void_date IS NULL
   AND aca.check_date BETWEEN :p_from_date AND :p_to_date
 ORDER BY fu.user_name, aca.check_date;


/* ---------------------------------------------------------------------
   Q5  Conflict exercised: user created a supplier AND entered its invoices
       (SUPPLIER_MAINT x AP_INVOICE_ENTRY)                    (validation)
   Swap asp.created_by for asp.last_updated_by to catch later edits.
   --------------------------------------------------------------------- */
SELECT fu.user_name            "Oracle User",
       asp.vendor_name         "Supplier",
       asp.creation_date       "Supplier Created",
       COUNT(aia.invoice_id)   "Invoices By Same User",
       SUM(aia.invoice_amount) "Invoice Value"
  FROM ap_suppliers    asp
  JOIN ap_invoices_all aia ON aia.vendor_id  = asp.vendor_id
                          AND aia.created_by = asp.created_by
  JOIN fnd_user        fu  ON fu.user_id     = asp.created_by
 WHERE aia.creation_date BETWEEN :p_from_date AND :p_to_date
 GROUP BY fu.user_name, asp.vendor_name, asp.creation_date
 ORDER BY 5 DESC;
