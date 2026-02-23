(defun _csvEscape (s / t)
  (setq t (if (null s) "" (vl-princ-to-string s)))
  (if (wcmatch t "*[,\"\n\r]*")
    (strcat "\"" (vl-string-subst "\"\"" "\"" t) "\"")
    t
  )
)

(defun _writeLine (fh s) (write-line s fh))

(defun _ptStr (p)
  (if (and p (= (type p) 'list) (= (length p) 3))
    (strcat (rtos (nth 0 p) 2 6) ";" (rtos (nth 1 p) 2 6) ";" (rtos (nth 2 p) 2 6))
    ""
  )
)

(defun _get (code alist default)
  (if (assoc code alist) (cdr (assoc code alist)) default)
)

(defun _countDictAdd (dict key / pair)
  (if (setq pair (assoc key dict))
    (subst (cons key (1+ (cdr pair))) pair dict)
    (cons (cons key 1) dict)
  )
)

(defun _explodeAndCountNestedInserts (ename / before after e dict ed bn)
  (setq dict nil)
  (setq before (entlast))

  (command "_.UNDO" "_MARK")
  (command "_.EXPLODE" ename)
  (setq after (entlast))

  ;; walk newly created entities from entnext(before) until we hit 'after'
  (setq e (entnext before))
  (while e
    (setq ed (entget e))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq bn (cdr (assoc 2 ed)))
        (if bn (setq dict (_countDictAdd dict bn)))
      )
    )
    (if (= e after)
      (setq e nil)
      (setq e (entnext e))
    )
  )

  (command "_.UNDO" "_BACK")
  dict
)

(defun c:DWGEXTRACT ( / outA outB outC fhA fhB fhC olderr msBtr e ed
                      hdl bn lay ip rot sx sy sz dict pair insertedCount )

  ;; always close files on error
  (setq olderr *error*)
  (defun *error* (msg)
    (if fhA (close fhA))
    (if fhB (close fhB))
    (if fhC (close fhC))
    (setq *error* olderr)
    (princ (strcat "\nDWGEXTRACT ERROR: " msg "\n"))
    (princ)
  )

  (setq outA (strcat (getvar "DWGPREFIX") "placements.csv"))
  (setq outB (strcat (getvar "DWGPREFIX") "nested_expanded.csv"))
  (setq outC (strcat (getvar "DWGPREFIX") "dynamic_properties.csv"))

  (setq fhA (open outA "w"))
  (setq fhB (open outB "w"))
  (setq fhC (open outC "w"))

  (_writeLine fhA "source_dwg,handle,block_name,layer,ins_pt_xyz,rotation_rad,scale_xyz,is_anonymous")
  (_writeLine fhB "source_dwg,outer_handle,outer_block_name,outer_layer,nested_block_name,nested_count")
  (_writeLine fhC "source_dwg,note")
  (_writeLine fhC (strcat (_csvEscape (getvar "DWGNAME")) ","
                          (_csvEscape "Dynamic props not exported (COM not used). Counts are from explode + nested INSERT scan.")))

  ;; Get Model Space block table record and iterate entities via entnext
  (setq msBtr (tblobjname "BLOCK" "*MODEL_SPACE"))
  (if (null msBtr)
    (progn
      (close fhA) (close fhB) (close fhC)
      (setq *error* olderr)
      (princ "\nERROR: Could not access *MODEL_SPACE block record.\n")
      (princ)
      (exit)
    )
  )

  (setq insertedCount 0)
  (setq e (entnext msBtr))
  (while e
    (setq ed (entget e))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq insertedCount (1+ insertedCount))

        (setq hdl (_get 5 ed ""))
        (setq bn  (_get 2 ed ""))
        (setq lay (_get 8 ed ""))
        (setq ip  (_get 10 ed '(0.0 0.0 0.0)))
        (setq rot (_get 50 ed 0.0))
        (setq sx  (_get 41 ed 1.0))
        (setq sy  (_get 42 ed 1.0))
        (setq sz  (_get 43 ed 1.0))

        (_writeLine fhA
          (strcat
            (_csvEscape (getvar "DWGNAME")) ","
            (_csvEscape hdl) ","
            (_csvEscape bn) ","
            (_csvEscape lay) ","
            (_csvEscape (_ptStr ip)) ","
            (_csvEscape (rtos rot 2 8)) ","
            (_csvEscape (strcat (rtos sx 2 6) ";" (rtos sy 2 6) ";" (rtos sz 2 6))) ","
            (_csvEscape (if (wcmatch (strcase bn) "`**U*") "true" "false"))
          )
        )

        ;; explode + count nested inserts
        (setq dict (_explodeAndCountNestedInserts e))
        (foreach pair dict
          (_writeLine fhB
            (strcat
              (_csvEscape (getvar "DWGNAME")) ","
              (_csvEscape hdl) ","
              (_csvEscape bn) ","
              (_csvEscape lay) ","
              (_csvEscape (car pair)) ","
              (_csvEscape (itoa (cdr pair)))
            )
          )
        )
      )
    )
    (setq e (entnext e))
  )

  (_writeLine fhC (strcat (_csvEscape (getvar "DWGNAME")) ","
                          (_csvEscape (strcat "Top-level INSERTs scanned in *MODEL_SPACE*: " (itoa insertedCount)))))

  (close fhA) (close fhB) (close fhC)
  (setq *error* olderr)

  (princ (strcat "\nWrote: " outA))
  (princ (strcat "\nWrote: " outB))
  (princ (strcat "\nWrote: " outC))
  (princ)
)
