;;; Command: TLEN
;;; Description: Calculates the total length of selected lines and polylines.

(defun c:TLEN (/ ss i ent tot_len len)
  (setq tot_len 0.0)
  (prompt "\nSelect objects to calculate total length (Lines, Polylines): ")
  (if (setq ss (ssget '((0 . "LINE,LWPOLYLINE,POLYLINE"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        
        ;; Retrieve the length property directly using AutoCAD core function (AutoCAD 2012+)
        ;; This does not require vl-load-com or VLAX extensions
        (setq len (vl-catch-all-apply 'getpropertyvalue (list ent "Length")))
        
        (if (and len (not (vl-catch-all-error-p len)))
          (setq tot_len (+ tot_len len))
        )
        
        (setq i (1+ i))
      )
      (alert (strcat "Total Length: " (rtos tot_len 2 4)))
      (princ (strcat "\nTotal Length: " (rtos tot_len 2 4)))
    )
    (princ "\nNo valid objects selected.")
  )
  (princ)
)
(princ "\nTotal Length LISP loaded. Type TLEN to run the command.")
(princ)
