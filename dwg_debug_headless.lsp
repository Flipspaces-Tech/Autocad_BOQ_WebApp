(princ "\n=== LOADED: dwg_debug_headless.lsp ===\n")

(defun _w (fh s) (write-line s fh))

(defun c:DWGDEBUG ( / out fh msBtr e ed i )
  (setq out (strcat (getvar "DWGPREFIX") "dwg_debug.txt"))
  (setq fh (open out "w"))

  (_w fh (strcat "DWGNAME=" (getvar "DWGNAME")))
  (_w fh (strcat "DWGPREFIX=" (getvar "DWGPREFIX")))
  (_w fh "STEP: tblobjname BLOCK *MODEL_SPACE*")

  (setq msBtr (tblobjname "BLOCK" "*MODEL_SPACE*"))
  (_w fh (strcat "msBtr=" (vl-princ-to-string msBtr)))
  (_w fh (strcat "type(msBtr)=" (vl-princ-to-string (type msBtr))))

  ;; Try both model space names (some files differ)
  (if (or (null msBtr) (= msBtr T))
    (progn
      (_w fh "msBtr nil/T. Trying *Paper_Space* just to check...")
      (setq msBtr (tblobjname "BLOCK" "*PAPER_SPACE*"))
      (_w fh (strcat "paperBtr=" (vl-princ-to-string msBtr)))
      (_w fh (strcat "type(paperBtr)=" (vl-princ-to-string (type msBtr))))
    )
  )

  (_w fh "STEP: entnext(msBtr)")
  (setq e (entnext msBtr))
  (_w fh (strcat "first e=" (vl-princ-to-string e)))
  (_w fh (strcat "type(e)=" (vl-princ-to-string (type e))))

  (setq i 0)
  (while (and e (< i 30))
    (setq ed (entget e))
    (if (null ed)
      (_w fh (strcat "entget nil at i=" (itoa i) " e=" (vl-princ-to-string e)))
      (_w fh (strcat "i=" (itoa i) " dxftype=" (cdr (assoc 0 ed))))
    )
    (setq e (entnext e))
    (setq i (1+ i))
  )

  (_w fh "DONE")
  (close fh)
  (princ (strcat "\nWrote debug: " out "\n"))
  (princ)
)
