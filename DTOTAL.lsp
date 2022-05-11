
(SETQ NOMD ""            ;DIRECTORI DATOS
      NOMDTIPO ""	        ;DIRECTORI DE TIPOS 	
      NOMDTEXT ""	;DIRECTORI DE TEXTES DE TOTAL
      NOMDEXCEL ""	;DIRECTORI DE TEXTES DE EXEL
      tabuladors-tituls-excel  1
      
      NOMDWG (getvar "dwgname")	;NOMBRE DEL DIBUJO  CON EXTENSION
      PREFIX (GETVAR "DWGPREFIX")
      nomdwG (SUBSTR NOMDWG 1 (- (STRLEN NOMDWG) 4)) ;NOMBRE  DE DIBUJO POR OMISION
      nomdw (STRCAT NOMD  NOMDWG)    ;NOMBRE DE FICHERO DE DATOS 
      tipomision "expertline.xls"  ;FICHERO DE TIPO POR OMISION
      EXTENSIOTIPO "XLS"
      EXTENSIOEXCEL "XLS"
      EXTENSIOTEXT "TXT"
      NOMCAPASTOTALS "*";Nombre de las capas donde buscar en la opcion busca dibujo
)

 ;(setq datos nil)

(defun c:EXPERTLINE ()
  (setq FITXERDAT (if FITXERDAT FITXERDAT  tipomision))
  (IF  T;(null DATOS);(OR (= acaba 2)
    (cambiaDAT  FITXERDAT))
  (if datos (progn
	      (buscadatos datos)
  (setq dh (load_dialog "DTOTAL"))
      (if (and dh (new_dialog "DATOS" dh))
	(PROGN
	  (IF (NULL NOMDAT)
	     (setq NOMDAT NOMDWG))
	  
	 (SETQ  TOTSDATOS T)
	  (MOSTRA)

	  
	 (set_tile "NOMDAT" NOMDAT)
	 (setq  VLISTLINE "0")
	  (SET_tile "VLISTA"  VLISTLINE)
	 (set_tile "ALT" (nth 3 (nth (atoi VLISTLINE) datos)))
	 
	 (action_tile "ALT" "(cambialt $VALUE)")
	 (action_tile "NOMDAT" "(SETQ NOMDAT $VALUE)")
	 (action_tile
	    "FICHERO"
	    "(cambiaDAT  (GETFILED \"Nombre MODELO\" FITXERDAT EXTENSIOTIPO  8))(done_dialog 2)"
	  )
	  (action_tile "BUSCA" "(SETQ TOTSDATOS (IF TOTSDATOS NIL T))(MOSTRA)")
	  (action_tile "BORRAL" "(BORRAL $VALUE)")
	  (action_tile "VLISTA" "(EDITLIN $VALUE)")
	  (action_tile "SALVA4" "(SALVAEXCEL)")
	  
	   (action_tile "BLOCS" "(done_dialog 11)")
	


	  (setq acaba (start_dialog))

	  (unload_dialog dh)


	  
      ));dialeg
  ));datos
    (if (= acaba 11)
	  (c:buscablocs)
    )
   (if (= acaba 2)
	  (c:EXPERTLINE)
    )
)

(defun cambialt (altval)
   (setq altura  altval
	 datosbb nil
	  ntag -1
	 )
    (foreach add datos
      (setq datosbb
	     (cons (if (= (setq ntag (1+ ntag)) (atoi VLISTLINE))
		     (list (car add) (nth 1 add) (nth 2 add) altura (nth 4 add))
		     add
		   )
		   datosbb
	     )
      )
      (setq datos (reverse datosbb))
      (MOSTRA)
    )
)

  
(DEFUN EDITLIN (VALIST)
     (SETQ  VLISTLINE VALIST
            ALTURA (NTH 3 (NTH (atoi VALIST) datos))
     )
     (set_tile "ALT" ALTURA)
)

(DEFUN BORRAL (NUMc)
  ;(set_tile "TEXT1" num)
  (SETQ DATOSA DATOS DATOS NIL
	DATOSQA DATOSQ DATOSQ NIL
	NNO -1
	NUM (if (null listamostra)(ATOI NUMc)
	                (cdr (assoc (ATOI numc) listamostra)))
  )
  (WHILE (SETQ AAA (NTH (SETQ NNO (1+ NNO)) DATOSA))
       (IF (/= NNO  NUM)
        (SETQ DATOS (CONS AAA DATOS)
	      DATOSQ (if datosqa (CONS (NTH NNO DATOSqA) DATOSQ)))
  ))
  (SETQ DATOS (REVERSE DATOS))
  (SETQ DATOSQ (REVERSE DATOSQ))

	    (MOSTRA)
  (SET_tile "VLISTA"  NUMc)
)






(DEFUN BUSCAdatos (LISTA)
   ;(if (scsemaf)(quit))
  (setq datosQ nil listatrib nil)
  (contablocs)
  (foreach datl lista (progn
         (SETQ LLL 0)
	 (if (< 2 (length datl))(progn
	 (setq TIPOF (CADR DATL)
	       CAPAS (CAR DATL)
	       NOMB (NTH 2 DATL)
	       %%% 1;(ATOI (NTH 3 DATL))
	  )

	 (set_tile "TEXT1" (STRCAT "Buscando y sumando datos de " NOMB))
         (cond ((= "BLQ" TIPOF)
		(SETQ  LLL   (SUMABLQ capas NOMB )) )
	       ((= "VIS" TIPOF)
		(SETQ  LLL   (SUMAvis capas NOMB )) )
	       ((= "ATR" TIPOF)
		(if (or (null listatrib)(null (setq listatrb1 (assoc capas listatrib))))
		  (SETQ listatrb1 (sumatrib (ssget "X" (list (cons 8  CAPAS)(CONS 66 1) )) )
		        listatrib (cons (cons capas (list listatrb1)) listatrib)
			listatrb1 (car  listatrib)
	          )
		)
	        (setq lll (if (and  listatrb1 (setq lll (assoc nomb (cadr listatrb1)))) (cdr lll) 0))
	       )
	       ((= "ATS" TIPOF)
		(SETQ atss (ssget "X"  (list (CONS 0 "INSERT") (cons 8  CAPAS) ;(CONS 2  NOMB)
					     ))
		      LLL (if atss (sumatribid  atss nomb)   0 )))
	       ((= "ATC" TIPOF)
		(SETQ atss (ssget "X"  (list (CONS 0 "INSERT") (cons 8  CAPAS) ;(CONS 1  NOMB)
					     ))
		      LLL (if atss (sumatribnum  atss nomb)   0 ))) 
	       ((= "CLR" TIPOF)
		(SETQ LLL (MIDELIN (ssget "X"  (list (cons 8  CAPAS) (CONS 62 (ATOI NOMB)))))))
	       ((= "ALT" TIPOF)
		(SETQ LLL (MIDELIN (ssget "X"  (list (cons 8  CAPAS))))))
                ((= "LIN" TIPOF)
		(SETQ LLL (MIDELIN (ssget "X"  (list (cons 8  CAPAS))))))
	       ((= "ARE" TIPOF)
		(SETQ LLL (MIDEAREA (ssget "X"  (list (cons 8  CAPAS)(CONS 0 "HATCH"))))))
	       ((= "TLN" TIPOF)
		(SETQ LLL (MIDELIN (ssget "X"  (list (cons 8  CAPAS)(CONS 6 NOMB))))))
	       (T NIL)
        )
	;(IF (/= 1000 %%%)(SETQ LLL (/ (* LLL %%%) 1000))); converio  a metres
	))
	(SETQ DATOSQ (CONS LLL DATOSQ))
  ))
  (SETQ DATOSQ (REVERSE DATOSQ))
    
  ;;;;;;;;;(CAMBIADAT FITXERDAT) (SETQ LLL (sslength  (ssget "X"  (list (CONS 0 "ATTDEF") (cons 8  "15.ELÈCTRIC" ) (CONS 2  "24")))))

  (set_tile "TEXT1" "Busqueda terminada..." )
)












 (defun salvaEXCEL ( / valor)
   (IF DATOSQ (progn
   (setq dtt (PRINC (strcat PREFIX nomdEXCEL NOMDAT "."  EXTENSIOEXCEL))
         dt1 (open dtt "w")
	 nnn -1
	 nnxls 1
   )
   (IF DT1 (progn
  (WHILE (SETQ add (NTH (SETQ NNN (1+ NNN)) DATOS))
    (if (/= 0 (setq valor (NTH NNN DATOSQ)))
      (progn
       (princ (rtosc valor) dt1)(princ "\t" dt1);: VALOR
        (princ  (RTOSC (ATOF (nth 3 add))) dt1); altura
	(princ "\t" dt1)
	(princ (strcat	"=A" (itoa Nnxls) "*B" (itoa Nnxls));(RTOSC (* (atof (nth 3 add))  VALOR))
		DT1);MUTIPLICACIO
	(setq nnxls (1+ nnxls))
	(princ "\t" dt1)			    
	(princ (last add) dt1)(princ "\n" dt1);descripcio
    ))
   )
    (set_tile "TEXT1" (strcat "Fichero guardado "dtt))
    (close dt1)
  )
  (ALERT "Error al guardar el Fichero\nEs posible que se este editando en Exel"))
)))


;torna el numero separat per comes
(defun rtosc (valor)
   (setq valf (fix valor)
         vald (- valor valf)
   )
   (if (= 0 vald)(itoa valf)
	 (strcat (itoa valf)","(substr  (rtos vald 2 2) 3))
   )
)



 





    (defun LEEDATOSF (FITX / LISTA FIDI)


      (print fitx)
      (setq fidi (open fitx "r"))
      (WHILE (AND FIDI (setq linf (read-line fidi)))
	(SETQ CCC ""
	      CCN 0
	      CCM NIL
	)
	(WHILE
	  (/= "" (SETQ CCL (SUBSTR LINF (SETQ CCN (1+ CCN)) 1)))
	   (IF (= "\t" CCL)
	     (SETQ CCM (CONS CCC CCM)
		   CCC ""
	     )
	     (SETQ CCC (STRCAT CCC CCL))
	   )
	)
	(SETQ LISTA (CONS (REVERSE (CONS CCC CCM)) LISTA))
      )
      (IF FIDI
	(CLOSE FIDI)
      )
      (SETQ LISTA (REVERSE LISTA))


    )




(DEFUN cambiaDAT (FITX)
  (if (OR (NULL FITX) (null (setq fitxtt (findfile fitx))))
    (setq fitxtt (GETFILED "Nombre MODELO" FITXERDAT EXTENSIOTIPO 8))
  )
  (if fitxtt
    (SETQ DATOSQ NIL
	  FITXERDAT fitxtt
	   NNNA	  -1
	   DATOS  (LEEDATOSF fitxtt)
    )
  )
)








(DEFUN MOSTRA ()
     (if datos (progn
       (start_list "VLISTA")
       (SETQ NNNA -1
	     nnex -1
	     listamostra nil);Especificar el nombre del cuadro de lista

        (WHILE (SETQ add (NTH (SETQ NNNA (1+ NNNA)) DATOS))
	  (IF DATOSQ
	    (if	(OR TOTSDATOS (/= (NTH NNNA DATOSQ) 0))
	      (progn
		(setq listamostra
		       (cons (cons (SETQ NNex (1+ NNex)) nnna)
			     listamostra
		       )
		)
		(IF (nth 4 add)(ADD_LIST (STRLIST (STRCAT ""
				  (IF (NTH NNNA DATOSQ)
				    (RTOS (NTH NNNA DATOSQ) 2 2)
				    "0");quantitat
				  " x  "
				  (nth 3 add); altura
				  " = "
				  (IF (NTH NNNA DATOSQ)
				    (RTOS (* (atof (nth 3 add))
					  (NTH NNNA DATOSQ)) 2 2)
				    "0")
				  " ")
				  (nth 4 add);descripcio				  
			  ))
		)
	      )
	      (if (or (null (nth 4 add)) (= "" (nth 4 add)))
		(progn
		  (setq	listamostra
			 (cons (cons (SETQ NNex (1+ NNex)) nnna)
			       listamostra
			 )
		  )
		  (ADD_LIST
		    (STRCAT " <N " (itoa nnex) ">   " (nth 0 add))
		  )
		)
	      )
	    )
	    (IF (nth 4 add) (ADD_LIST (strcat " < "
			      (itoa nnna)
			      ">   "	(nth 0 add)  "   "
			        (nth 2 add); tipus
			      "   "
				 (nth 3 add); altura
			      "   "
				(nth 4 add)
			
			      
		      )
	    ))
	  )
	)
        (end_list)
       ))
)


;(strlist "22" "gggg" )

(DEFUN STRLIST (CAD1 CAD2)
  (SETQ LL (STRLEN CAD1)
	nn  (fix (* 1.65 (- 25 LL))))
  (repeat nn
    (SETQ CAD1 (STRCAT CAD1 " ") ))
  (STRCAT CAD1 CAD2)
)






(DEFUN sumatrib (OBC / listatrb-obc nno atr1 atr2)    
  (SETQ NNO -1  listatrb-obc nil)
   (IF OBC (PROGN
  (while (setq nob (ssname obc (setq nno (1+ nno))))
         (setq listatrb (attrs nob))
         (foreach atr1 listatrb (setq listatrb-obc (if  (and listatrb-obc
							     (setq atr2 (assoc (car atr1) listatrb-obc)))
				  			(subst (cons (car atr1)(+ (atof (cdr atr1)) (cdr atr2)))
							    atr2 listatrb-obc)
						        (cons (cons (car atr1) (atof (cdr atr1))) listatrb-obc)
						   )))
         )
  (reverse listatrb-obc)
     ))
)




(DEFUN sumatribnum (obc num /  nno atr1 atr2)   ; busca amb mateix id 
  (SETQ NNO -1  numcoincident 0 );(setq num "1")
   (IF OBC (PROGN
  (while (setq nob (ssname obc (setq nno (1+ nno))))

    (if	(and (setq listatrb (attrs nob))
	     (= num (cdr (car listatrb)))
	)
      (setq numcoincident (1+ numcoincident))
    )
  )
    (setq numcoincident  numcoincident)
  ))
)


(DEFUN sumatribid (OBC id / listatrb-obc nno atr1 atr2)   ; busca amb mateix id 
  (SETQ NNO -1  numcoincident 0  listatrb-obc nil)
   (IF OBC (PROGN
  (while (setq nob (ssname obc (setq nno (1+ nno))))
          (if (setq listatrb (attrs nob))
         (if (= id (car (car listatrb)))(setq numcoincident (1+ numcoincident)))
     ))
    (setq numcoincident  numcoincident)
  ))
)




 (DEFUN sumaLISTDAT (LISTA NOM CANT / ELE nno)    
  (SETQ NNO -1  )
  (SETQ LISTA (IF (AND LISTA (SETQ ELE (ASSOC NOM LISTA)))
      		(SUBST (CONS (CAR ELE) (+ (CDR ELE) CANT) ) ELE LISTA)
     	    	(cons (cons NOM CANT ) lista))
  )
 )

 (DEFUN SUMALISTAS (LISTA1 LISTA2)
         (foreach atr1 lista2 (setq lista1 (if  (and lista1 (setq atr2 (assoc (car atr1) lista1)))
				  			(subst (cons (car atr1)(+ (atof (cdr atr1)) (cdr atr2)))
							    atr2 lista1)
						        (cons (cons (car atr1) (atof (cdr atr1))) lista1)
						   )))
   (SETQ LISTA1 LISTA1)
)


    
(DEFUN  MIDELIN (OBC)
    
  (SETQ NNO -1 SST 0.0 sss 0)
  (if obc
  (while (setq nob (ssname obc (setq nno (1+ nno))))
         (setq nge (entget nob)
               nne (cdr (assoc 0 nge)))
        (cond ((= 1 (cdr (assoc 67  nge))) nil);es del espai paper
	      ((= nne "LINE") (setq sss  (distance (cdr (assoc 10 nge))
                                                      (cdr (assoc 11 nge)))
                                    ))
              ((= nne "ARC") (setq Anc (- (CDR (ASSOC 51 nge))
                                          (cdr (assoc 50 nge)))
                                   anc (if (minusp anc) (+ anc pi pi) anc)
                                   sss (* (cdr (assoc 40 nge)) anc)))
                                                
              ((= nne "INSERT") (if (setq lat (attrs nob))
                                    (if (assoc "L" lat)
                                                 (setq sss (ATOF (cdr 
                                                           (assoc "L" lat)))
                                )))
                               )
	      ((OR (= nne "POLYLINE")(= nne "LWPOLYLINE"))
                   (SETQ SSS (midepol nob nne))
	       )
	      (T NIL)
	  )
    (SETQ SST (+ SSS SST))
    )
  )
  (SETQ SST (/ SST  1000)); conversion a mm a metros
 )	     

(defun c:areasel ()
  (setq selare (ssget)
   )
  (midearea selare)
 )


(DEFUN  MIDEAREA (OBC)
    
  (SETQ NNO -1 SST 0.0 sss 0)
  (if obc
  (while (setq nob (ssname obc (setq nno (1+ nno))))
         (setq nge (entget nob)
               nne (cdr (assoc 0 nge)))
        (if (or (null (assoc 67 nge)) (= 0 (cdr (assoc 67 nge))))
	        (PROGN
	          (Command "AREA" "O" NOB)
                   (SETQ SSS (GETVAR "AREA"))
		)
          )
    (SETQ SST (+ SSS SST) )
    )
  )
  (SETQ SST   SST )
)

(DEFUN SUMABLQ (capa NOMBLQ)
   (SETQ	NNO -1
	SST 0.0
	sss 0
  ) ;(if  lst (progn

  (foreach cbk lst
    (if (and (eq capa (car cbk))(eq nomblq (cadr cbk)))
	(SETQ SSS (1+ SSS)
	)
    )
  )
  
  (SETQ SST SSS)
)

(DEFUN SUMAvis (capa NOMvis)
   (SETQ	NNO -1
	SST 0.0
	sss 0
  ) ;(if  lst (progn

  (foreach cbk lst
    (if (and (EQ capa (car Cbk))(EQ nomvis (CAR (last  cbk))))
	(SETQ SSS (1+ SSS)
	)
    )
  )
  
  (SETQ SST SSS)
)

	       
 (DEFUN MIDEPOL (PP1 TIPOPOL / PPA)	       
   (setq  ddd 0
          EJC NIL
	  ppg (entget pp1)
	  listap nil
     )
  (if (= "LWPOLYLINE" TIPOPOL)
 (foreach ppp ppg 
            (if (= 10 (car ppp))
             (setq ppt (cdr ppp)
		   DDD (IF PPA
			 (+ DDD (if (/= 0.0 p42)
			             (longarc ppa ppt p42)
			             (DISTANCE PPA PPT)))
		       DDD)		   
                   ppa ppt  
             )
           (IF (= 42 (car ppp))
	     (setq p42 (cdr ppp)
             )
	   ))
 )
 (progn
  (setq ppp (entnext pp1)   
  )
    (while (/= "SEQEND" (cdr (assoc 0 (setq ppg (entget ppp)))))
           (setq ppt (cdr (assoc 10 ppg))              
		 DDD (IF PPA
		       (+ DDD (if (/= 0.0 p42)
			         (longarc ppa ppt p42)
			         (DISTANCE PPA PPT)))
		      ddd )
                 ppa ppt
		 p42 (CDR (ASSOC 42 ppg))                    
                 ppp (entnext ppp)
           )
    )
  ))
 
  (if ddd ddd 0)
 )    


(DEFUN longarc (pt1 pt2 atn / alf pta pd1 pd2 pd3 pt3 ppc pdr prr pai paf pax pan plon)
(SETQ ALF (Atan atn)
      PTA (ANGLE PT1 PT2)
      PD1 (/ (DISTANCE PT1 PT2) 2)
      PD2 (* PD1 ATN) 
      PD3 (/ pd1 (cos alf))
      PT3 (POLAR PT1 PTA PD1)
      p90 (- PTA (/ PI 2))
      PT3 (POLAR PT3 p90 PD2)
      pdr (/ pd3 (* 2 (sin alf)))   ;radi
      ppc (polar pt3 (+ p90 pi) pdr);centre
      pdr (abs pdr)
       prr (= atn (abs atn))

       Pai  (if prr (angle ppc pt1)(angle ppc pt2))
       paf  (if prr (angle ppc pt2)(angle ppc pt1))

       pax (if (< pai paf) paf (+ paf pi pi))
       pan (if (< pai paf) pai pai)

       pad (- pax pan);angle
       plon (* pdr pad)
)
)





(defun ATTRs (sss / ssv snx)
  (setq snx (entnext sss)
        SSV NIL)
  (while snx
     (setq ssg (entget snx))
     (if (= "ATTRIB" (cdr (assoc 0 ssg)))     
             (setq ssv (cons (cons (cdr (assoc 2 ssg)) 
                                   (cdr (assoc 1 ssg))) ssv) 
                   snx (entnext snx))
             (setq snx nil)
     ) 
  )
  (if ssv ssv)
)

(defun LM:blockname ( obj )
    (if (vlax-property-available-p obj 'effectivename)
        (defun LM:blockname ( obj ) (vla-get-effectivename obj))
        (defun LM:blockname ( obj ) (vla-get-name obj))
    )
    (LM:blockname obj)
)

(defun LM:getvisibilityparametername ( blk / vis )  
    (if
        (and
            (vlax-property-available-p blk 'effectivename)
            (setq blk
                (vla-item
                    (vla-get-blocks (vla-get-document blk))
                    (vla-get-effectivename blk)
                )
            )
           ;(= :vlax-true (vla-get-isdynamicblock blk)) to account for NUS dynamic blocks
            (= :vlax-true (vla-get-hasextensiondictionary blk))
            (setq vis
                (vl-some
                   '(lambda ( pair )
                        (if
                            (and
                                (= 360 (car pair))
                                (= "BLOCKVISIBILITYPARAMETER" (cdr (assoc 0 (entget (cdr pair)))))
                            )
                            (cdr pair)
                        )
                    )
                    (dictsearch
                        (vlax-vla-object->ename (vla-getextensiondictionary blk))
                        "ACAD_ENHANCEDBLOCK"
                    )
                )
            )
        )
        (cdr (assoc 301 (entget vis)))
    )
)

(c:expertline)