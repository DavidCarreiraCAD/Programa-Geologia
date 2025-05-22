(defun c:LogTransformMultiple ( / )

  ;; ============================================
  ;; BLOCO 1 - INPUT DE ESCALA E CÁLCULO DE FATOR
  ;; ============================================

  (vl-load-com)

;; ALTERAR AQUI!!!

;; Introduzir escala HORIZONTAL

(setq escala XXXXXX)

  (setq fator (/ escala 3000.0))
  (princ (strcat "\n🔢 Fator de escala aplicado: " (rtos fator 2 6)))

  ;; ===============================
  ;; BLOCO 2 - Formatação dos textos
  ;; ===============================

  (setq style (tblsearch "style" "Standard"))
  (if style
    (progn
      (command "-style" "Standard" "Arial" "" "" "" "" "" "")
      (princ "\nFonte do estilo 'Standard' alterada para Arial.")
    )
    (princ "\nEstilo 'Standard' não encontrado.")
  )

  (setq ss (ssget "X" '((0 . "TEXT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq entData (entget ent))
        (setq content (strcase (cdr (assoc 1 entData))))
        (setq height (cdr (assoc 40 entData)))

        ;; Mover e rodar certos textos
        (setq moveY
          (cond
            ((= content "NSPT") (* 6.0 fator))
            ((= content "LU/LF") (* 7.7 fator))
            ((= content "REC/RQD") (* 12.0 fator))
            (T nil)
          )
        )

        (if moveY
          (progn
            (setq entData (subst (cons 50 (/ pi 2)) (assoc 50 entData) entData))
            (setq alignPt (assoc 11 entData))
            (if alignPt
              (setq entData (subst
                              (cons 11 (list (car (cdr alignPt))
                                            (+ (cadr (cdr alignPt)) moveY)
                                            (caddr (cdr alignPt))))
                              alignPt entData))
              (progn
                (setq insPt (assoc 10 entData))
                (setq entData (subst
                                (cons 10 (list (car (cdr insPt))
                                               (+ (cadr (cdr insPt)) moveY)
                                               (caddr (cdr insPt))))
                                insPt entData))
              )
            )
          )
        )

        (if (= height 0.1368)
  (progn
    ;; Aplica rotação vertical (90º)
    (setq entData (subst (cons 50 (/ pi 2)) (assoc 50 entData) entData))

    ;; Aplica nova altura com fator de escala
    (setq entData (subst (cons 40 (* 2.5 fator)) (assoc 40 entData) entData))

    ;; Define ponto de inserção
    (setq insPt (cdr (assoc 10 entData)))

    ;; Garante que o ponto de alinhamento (11) existe e é igual ao ponto de inserção
    (if (assoc 11 entData)
      (setq entData (subst (cons 11 insPt) (assoc 11 entData) entData))
      (setq entData (append entData (list (cons 11 insPt))))
    )

    ;; Define justificação horizontal (72) para Right
    (if (assoc 72 entData)
      (setq entData (subst (cons 72 2) (assoc 72 entData) entData))
      (setq entData (append entData (list (cons 72 2))))
    )

    ;; Define justificação vertical (73) para Top
    (if (assoc 73 entData)
      (setq entData (subst (cons 73 3) (assoc 73 entData) entData))
      (setq entData (append entData (list (cons 73 3))))
    )
  )
)
        (if (= height 0.144)
          (setq entData (subst (cons 40 (* 2.5 fator)) (assoc 40 entData) entData))
        )

        (if (= height 3.996)
          (setq entData (subst (cons 40 (* 3.996 fator)) (assoc 40 entData) entData))
        )

        (entmod entData)
        (setq i (1+ i))
      )
      (princ "\n📄 Formatação de textos concluída.")
    )
    (princ "\nNenhum texto encontrado.")
  )

  ;; ================================
  ;; BLOCO 3 - Hatch por layer "Log_"
  ;; ================================

  (setq acadApp (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acadApp))
  (setq ms (vla-get-ModelSpace doc))

  (setq ssAll (ssget "_X" '((0 . "*POLYLINE"))))

  (if (not ssAll)
    (progn (princ "\n❌ Nenhuma polyline encontrada no desenho.") (exit))
  )

  (setq logLayers '())
  (setq i 0)
  (while (< i (sslength ssAll))
    (setq ent (ssname ssAll i))
    (setq entObj (vlax-ename->vla-object ent))
    (setq lyr (vla-get-Layer entObj))
    (if (and (wcmatch lyr "Log_*") (not (member lyr logLayers)))
      (setq logLayers (cons lyr logLayers))
    )
    (setq i (1+ i))
  )

  (if (null logLayers)
    (progn (princ "\n⚠ Nenhuma layer com prefixo Log_ encontrada.") (exit))
  )

  (setq width-addition-alist '(("PZ" . 2.592)
                               ("W" . 7.2)
                               ("F" . 7.2)
                               ("REC/RQD" . 10.368)
                               ("LU/LF" . 4.32)
                               ("NSPT" . 10.368)))

  (foreach lyr logLayers
    (princ (strcat "\n🔍 Processando layer: " lyr))

    (setq ssP (ssget "_X" (list (cons 0 "*POLYLINE") (cons 8 lyr))))
    (if (not ssP)
      (princ "\n⚠ Nenhuma polyline nesta layer.")
      (progn
        (setq maxX -1e99 minY 1e99 maxHeight 0 iP 0)

        (while (< iP (sslength ssP))
          (setq ename (ssname ssP iP))
          (setq entP (vlax-ename->vla-object ename))
          (vla-GetBoundingBox entP 'minPt 'maxPt)
          (setq minPt (vlax-safearray->list minPt))
          (setq maxPt (vlax-safearray->list maxPt))

          (if (> (car maxPt) maxX) (setq maxX (car maxPt)))
          (if (< (cadr minPt) minY) (setq minY (cadr minPt)))

          (setq thisHeight (- (cadr maxPt) (cadr minPt)))
          (if (> thisHeight maxHeight) (setq maxHeight thisHeight))

          (setq iP (1+ iP))
        )

        (setq largura (* 4.1774 fator))
        (setq deslocamento-fixo (* -4.1774 fator))
        (setq soma-deducoes 0.0)

        (setq ssText (ssget "_X" (list (cons 0 "TEXT") (cons 8 lyr))))

        (if ssText
          (progn
            (setq iText 0)
            (while (< iText (sslength ssText))
              (setq textEnt (ssname ssText iText))
              (setq textObj (vlax-ename->vla-object textEnt))
              (setq rawText (vla-get-TextString textObj))
              (setq textCont (strcase (vl-string-trim " " rawText)))

              (foreach pair width-addition-alist
                (if (= textCont (car pair))
                  (progn
                    (setq largura (+ largura (* (cdr pair) fator)))
                    (setq soma-deducoes (+ soma-deducoes (* (cdr pair) fator)))
                    (princ (strcat "\n➕ Somado " (rtos (* (cdr pair) fator) 2 3) " por conteúdo: " textCont))
                  )
                )
              )
              (setq iText (1+ iText))
            )
          )
        )

        (setq pt1 (list maxX minY))
        (setq pt2 (list (+ maxX largura) minY))
        (setq pt3 (list (+ maxX largura) (+ minY maxHeight)))
        (setq pt4 (list maxX (+ minY maxHeight)))
        (setq coords (apply 'append (list pt1 pt2 pt3 pt4 pt1)))

        (setq safearray (vlax-make-safearray vlax-vbDouble (cons 0 (- (length coords) 1))))
        (vlax-safearray-fill safearray coords)
        (setq rect (vla-AddLightWeightPolyline ms (vlax-make-variant safearray)))
        (vla-put-Closed rect :vlax-true)

        (setq hatchObj (vla-AddHatch ms acHatchPatternTypePreDefined "SOLID" :vlax-true))
        (vla-put-Layer hatchObj lyr)
        (setq loopArray (vlax-make-safearray vlax-vbObject '(0 . 0)))
        (vlax-safearray-put-element loopArray 0 rect)
        (setq loopVariant (vlax-make-variant loopArray))
        (vla-AppendOuterLoop hatchObj loopVariant)
        (vla-put-Color hatchObj 254)
        (vla-Evaluate hatchObj)

        ;; Mover hatch com base em deslocamento fixo menos as deduções
        (command "_.draworder" (vlax-vla-object->ename hatchObj) "" "back")
        (setq dx (- deslocamento-fixo soma-deducoes))
        (vla-Move hatchObj (vlax-3d-point (list 0 0 0)) (vlax-3d-point (list dx 0 0)))

        (vla-Delete rect)

        (princ "\n✅ Hatch criado com sucesso.")
      )
    )
  )

(setq ssLog (ssget "_X" '((8 . "Log_*"))))
(if ssLog
  (command "_.draworder" ssLog "" "front")
)

  ;; ===================================
  ;; BLOCO 4 - Importa e substitui prosp
  ;; ===================================

(vl-load-com)
(setq ficheiro "C:\\Users\\DPC\\OneDrive - COBAGroup\\Desktop\\REVIT_David\\SONDAGENS_DWG_TESTES\\ZZZ_ProspBlocks.dwg")
(setq acadApp (vlax-get-acad-object))
(setq thisDoc (vla-get-ActiveDocument acadApp))

;; === IMPORTAR BLOCOS ===
(princ "\n[1/2] A importar blocos...")

(setq sourceDoc (vla-Open (vla-get-Documents acadApp) ficheiro :vlax-false))
(setq blocks (vla-get-Blocks sourceDoc))

(vlax-for blk blocks
  (if (and (not (wcmatch (vla-get-Name blk) "*Model_Space, *Paper_Space"))
           (/= (vla-get-IsXRef blk) :vlax-true)
           (/= (vla-get-IsLayout blk) :vlax-true))
    (vla-CopyObjects sourceDoc
      (vlax-make-variant
        (vlax-safearray-fill
          (vlax-make-safearray vlax-vbObject (cons 0 0))
          (list blk)))
      (vla-get-Blocks thisDoc))
  )
)

(vla-Close sourceDoc :vlax-false)
(vla-Regen thisDoc acAllViewports)
(princ "\nBlocos importados com sucesso.")

;; ========================== CSV HELPERS ==========================
(defun split-string (str delim / lst tmp ch)
  (setq lst '() tmp "")
  (foreach ch (vl-string->list str)
    (if (/= ch (ascii delim))
      (setq tmp (strcat tmp (chr ch)))
      (progn
        (setq lst (append lst (list tmp)))
        (setq tmp "")
      )
    )
  )
  (setq lst (append lst (list tmp)))
  lst
)

(defun ler-correspondencias ( / file line fields nome tipo cor bloco correspondencias)
  (setq file (open "C:\\Users\\DPC\\OneDrive - COBAGroup\\Desktop\\REVIT_David\\SONDAGENS_DWG_TESTES\\GT_RW-Exemplo.csv" "r"))
  (setq correspondencias '())
  (if file
    (progn
      (read-line file) ; ignora cabeçalho
      (while (setq line (read-line file))
        (setq fields (split-string line ";"))
        (if (>= (length fields) 3)
          (progn
            (setq nome (nth 0 fields))
            (setq tipo (nth (- (length fields) 2) fields))
            (setq cor (nth (- (length fields) 1) fields))
            (setq bloco (strcat tipo "_" cor))
            (setq correspondencias (cons (cons nome bloco) correspondencias))
          )
        )
      )
      (close file)
    )
  )
  correspondencias
)

;; ========================== PRINCIPAL ==========================
(setq correspondencias (ler-correspondencias))
(setq ss (ssget "X" '((0 . "TEXT") (40 . 3.6))))

(setvar "ATTDIA" 0)
(setvar "ATTREQ" 0)

(if ss
  (progn
    (setq i 0)
    (while (< i (sslength ss))
      (setq ent (ssname ss i))
      (setq eData (entget ent))
      (setq nome (cdr (assoc 1 eData)))
      (setq bloco (cdr (assoc nome correspondencias)))
      (if bloco
        (progn
          (setq pt (cdr (assoc 10 eData)))
          (setq x (car pt))
          (setq y (cadr pt))
          (setq z (caddr pt))

          ;; Ajuste por tipo
          (setq x (- x 0.1607))
          (if (wcmatch nome "P*")
            (setq y (- y 3.4393)) ; deslocamento para P
            (setq y (- y 19.5673)) ; default deslocamento S
          )

          (setq newPt (list x y z))
(setq layer (cdr (assoc 8 eData))) ; obter layer do texto
(entdel ent)

;; Corrigir fator para compensar bloco com escala interna 12
(setq fatorCorrigido (* fator 12))

;; Inserir o bloco
(command "_.-INSERT" bloco newPt fatorCorrigido fatorCorrigido 0)

;; Definir layer do bloco para a mesma do texto
(setq blkEnt (entlast))
(if blkEnt
  (entmod (subst (cons 8 layer) (assoc 8 (entget blkEnt)) (entget blkEnt)))
)


          ;; Atualizar atributos
          (setq blkEnt (entlast))
          (if (and blkEnt (tblsearch "BLOCK" bloco))
            (progn
              (setq blkData (entnext blkEnt))
              (while (and blkData (/= (cdr (assoc 0 (entget blkData))) "SEQEND"))
                (if (= (cdr (assoc 0 (entget blkData))) "ATTRIB")
                  (progn
                    (setq attData (entget blkData))
                    (entmod (subst (cons 1 nome) (assoc 1 attData) attData))
                    (entupd blkData)
                  )
                )
                (setq blkData (entnext blkData))
              )
            )
          )
        )
      )
      (setq i (1+ i))
    )
  )
  (princ "\nNenhum texto encontrado com height 3.6.")
)

(setvar "ATTDIA" 1)
(setvar "ATTREQ" 1)

(princ "\n✅ Substituição concluída com escala e move adaptados.")
(princ)

   ;; ===========================================
  ;; BLOCO 5 - Salvar e fechar o arquivo DWG
  ;; ===========================================

   (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))) ; Obtém o documento ativo
  (setq path (getvar "DWGPREFIX")) ; Caminho da pasta onde o desenho está
  (setq name (vl-filename-base (getvar "DWGNAME"))) ; Nome do arquivo sem extensão
  (setq newname (strcat path name))

  ;; Salva como DWG 2018
  (vla-saveas doc newname ac2018) ; "ac2018" é o código para salvar como DWG 2018

  (princ (strcat "\nFicheiro guardado como: " newname))
  (command "_quit")
  (princ)
)

