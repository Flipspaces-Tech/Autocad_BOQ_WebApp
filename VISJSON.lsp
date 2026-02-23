(vl-load-com)

;; =========================
;; CONFIG
;; =========================
(setq *VISJSON_EXPORT_DIR* "C:/Users/admin/Documents/AUTOCAD_WEBAPP/EXPORTS")
(setq *VISJSON_JSON_FILE*  "vis_export.json")
(setq *VISJSON_LOG_FILE*   "vis_export_log.txt")

;; =========================
;; helpers
;; =========================
(defun _obj? (o) (and o (vlax-objectp o)))
(defun _str  (v) (if v (vl-princ-to-string v) ""))
(defun _upper(s) (if s (strcase (_str s)) ""))

(defun _ensure-dir (p)
  (if (not (vl-file-directory-p p)) (vl-mkdir p))
  p
)

(defun _log (msg / f fn)
  (_ensure-dir *VISJSON_EXPORT_DIR*)
  (setq fn (strcat *VISJSON_EXPORT_DIR* "/" *VISJSON_LOG_FILE*))
  (setq f (open fn "a"))
  (write-line msg f)
  (close f)
)

(defun _json-esc (s / out i ch)
  (setq s (_str s))
  (setq out "" i 1)
  (while (<= i (strlen s))
    (setq ch (substr s i 1))
    (cond
      ((= ch "\\") (setq out (strcat out "\\\\")))
      ((= ch "\"") (setq out (strcat out "\\\"")))
      ((= ch "\n") (setq out (strcat out "\\n")))
      ((= ch "\r") (setq out (strcat out "\\r")))
      ((= ch "\t") (setq out (strcat out "\\t")))
      (T (setq out (strcat out ch)))
    )
    (setq i (1+ i))
  )
  out
)

(defun _safe->list (x)
  (cond
    ((= (type x) 'VARIANT) (_safe->list (vlax-variant-value x)))
    ((= (type x) 'SAFEARRAY) (vlax-safearray->list x))
    ((= (type x) 'LIST) x)
    (T nil)
  )
)

(defun _get (obj prop / r)
  (setq r (vl-catch-all-apply 'vlax-get (list obj prop)))
  (if (vl-catch-all-error-p r) nil r)
)

(defun _inv (obj method / r)
  (setq r (vl-catch-all-apply 'vlax-invoke (list obj method)))
  (if (vl-catch-all-error-p r) nil r)
)

(defun _num? (s / x)
  (setq s (vl-string-trim " \t\r\n" (_str s)))
  (if (= s "") nil
    (progn (setq x (distof s 2)) (if x T nil))
  )
)

;; =========================
;; dynamic block info
;; =========================
(defun _effective-name (vlaBlk / n)
  (setq n (_get vlaBlk 'EffectiveName))
  (if (null n) (setq n (_get vlaBlk 'Name)))
  (_str n)
)

(defun _get-visibility (vlaBlk / var props p pName pVal vis)
  (setq vis "")
  (if (and (_obj? vlaBlk) (= :vlax-true (_get vlaBlk 'IsDynamicBlock)))
    (progn
      (setq var (_inv vlaBlk 'GetDynamicBlockProperties))
      (setq props (_safe->list var))
      (if props
        (foreach p props
          (if (_obj? p)
            (progn
              (setq pName (_upper (_get p 'PropertyName)))
              (if (wcmatch pName "*VISIBILITY*")
                (progn
                  (setq pVal (_get p 'Value))
                  (setq vis (_str pVal))
                )
              )
            )
          )
        )
      )
    )
  )
  vis
)

(defun _get-chaircount-from-dynprops (vlaBlk / var props p pName pVal val out)
  ;; returns integer or nil
  (setq out nil)
  (if (and (_obj? vlaBlk) (= :vlax-true (_get vlaBlk 'IsDynamicBlock)))
    (progn
      (setq var (_inv vlaBlk 'GetDynamicBlockProperties))
      (setq props (_safe->list var))
      (if props
        (foreach p props
          (if (and (_obj? p) (null out))
            (progn
              (setq pName (_upper (_get p 'PropertyName)))
              (if (or (wcmatch pName "*CHAIR*")
                      (wcmatch pName "*SEAT*")
                      (wcmatch pName "*COUNT*")
                      (wcmatch pName "*QTY*")
                      (wcmatch pName "*NO*CHAIR*")
                      (wcmatch pName "*NO*SEAT*"))
                (progn
                  (setq pVal (_get p 'Value))
                  (setq val (_str pVal))
                  (if (_num? val)
                    (setq out (fix (+ 0.5 (distof val 2))))
                  )
                )
              )
            )
          )
        )
      )
    )
  )
  out
)

;; =========================
;; grouping map: key -> (table_count . chair_sum)
;; =========================
(defun _map-get (m k) (cdr (assoc k m)))

(defun _map-set (m k table_add chair_add / cur t c)
  (setq cur (_map-get m k))
  (if cur
    (progn
      (setq t (+ (car cur) table_add))
      (setq c (+ (cdr cur) chair_add))
      (subst (cons k (cons t c)) (assoc k m) m)
    )
    (cons (cons k (cons table_add chair_add)) m)
  )
)

;; =========================
;; build json
;; =========================
(defun _build-json (m / out first p k v sep name vis tcount csum)
  (setq out "{\n  \"items\": [\n")
  (setq first T)

  (foreach p m
    (setq k (car p))
    (setq v (cdr p))
    (setq tcount (car v))
    (setq csum   (cdr v))

    (setq sep (vl-string-search "|" k))
    (setq name (if sep (substr k 1 sep) k))
    (setq vis  (if sep (substr k (+ sep 2)) ""))

    (if first (setq first nil) (setq out (strcat out ",\n")))

    (setq out
      (strcat out
        "    {"
        "\"name\":\"" (_json-esc name) "\", "
        "\"visibility\":\"" (_json-esc vis) "\", "
        "\"table_count\":" (itoa (max 0 (fix tcount))) ", "
        "\"chairs\":" (itoa (max 0 (fix csum)))
        "}"
      )
    )
  )

  (setq out (strcat out "\n  ]\n}\n"))
  out
)

(defun _write-json (txt / fn f)
  (_ensure-dir *VISJSON_EXPORT_DIR*)
  (setq fn (strcat *VISJSON_EXPORT_DIR* "/" *VISJSON_JSON_FILE*))
  (setq f (open fn "w"))
  (write-line txt f)
  (close f)
  fn
)

;; =========================
;; get all INSERTs without doc/modelspace
;; =========================
(defun _collect-inserts-ss (/ ss i e lst)
  (setq lst '())
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq lst (cons e lst))
        (setq i (1+ i))
      )
      (reverse lst)
    )
    nil
  )
)

;; =========================
;; MAIN
;; =========================
(defun c:VISJSON (/ f blocks m en vla name vis chairs k json fn nDyn)
  ;; reset log
  (_ensure-dir *VISJSON_EXPORT_DIR*)
  (setq f (open (strcat *VISJSON_EXPORT_DIR* "/" *VISJSON_LOG_FILE*) "w"))
  (write-line "[VISJSON] start" f)
  (write-line (strcat "[VISJSON] DWGNAME=" (getvar "DWGNAME")) f)
  (close f)

  (setq blocks (_collect-inserts-ss))
  (setq m '())
  (setq nDyn 0)

  (if (null blocks)
    (progn
      (_log "[VISJSON] no INSERT entities found. writing empty json.")
      (setq json (_build-json '()))
      (setq fn (_write-json json))
      (_log (strcat "[VISJSON] wrote: " fn))
      (prompt (strcat "\n✅ JSON written to: " fn))
      (princ)
    )
    (progn
      (_log (strcat "[VISJSON] total INSERTs: " (itoa (length blocks))))

      (foreach en blocks
        (setq vla (vl-catch-all-apply 'vlax-ename->vla-object (list en)))
        (if (and (not (vl-catch-all-error-p vla)) (_obj? vla))
          (if (= :vlax-true (_get vla 'IsDynamicBlock))
            (progn
              (setq nDyn (1+ nDyn))
              (setq name (_effective-name vla))
              (setq vis  (_get-visibility vla))

              ;; chair count: numeric dynprop only (headless-safe)
              (setq chairs (_get-chaircount-from-dynprops vla))
              (if (null chairs) (setq chairs 0))

              ;; table_count = 1 per dynamic block insert
              (setq k (strcat name "|" vis))
              (setq m (_map-set m k 1 chairs))
            )
          )
        )
      )

      (_log (strcat "[VISJSON] dynamic INSERTs processed: " (itoa nDyn)))

      (setq m (reverse m))
      (setq json (_build-json m))
      (setq fn (_write-json json))
      (_log (strcat "[VISJSON] done. wrote: " fn))

      (prompt (strcat "\n✅ JSON written to: " fn))
      (princ)
    )
  )
)

(princ "\nVISJSON loaded. Run command: VISJSON\n")
(princ)
