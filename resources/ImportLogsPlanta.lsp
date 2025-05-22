(defun c:ImportLogsPlanta ( / ficheiroCsv ficheiroBlocos acadApp thisDoc sourceDoc blocks dados fatorCorrigido)

  ;; === CONFIGURAÇÕES ===
  (setq ficheiroBlocos "C:/Users/DPC/OneDrive - COBAGroup/Desktop/DPC/AutoLisp/ImportLogsPlanta/ProspBlocks.dwg")
  (setq ficheiroCsv    "C:/Users/DPC/OneDrive - COBAGroup/Desktop/DPC/AutoLisp/ImportLogsPlanta/GT_RW-Exemplo.csv")

  (princ "\n[1/3] A importar blocos...")

  (vl-load-com)
  (setq acadApp (vlax-get-acad-object))
  (setq thisDoc (vla-get-ActiveDocument acadApp))

  ;; === ALTERAR UNIDADES PARA METROS ===
  (setvar "INSUNITS" 6)

  (setq sourceDoc (vla-Open (vla-get-Documents acadApp) ficheiroBlocos :vlax-false))
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

  ;; === FUNÇÃO PARA DIVIDIR STRINGS ===
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

  ;; === FUNÇÃO PARA EXTRAIR LETRAS DO INÍCIO DO CÓDIGO ===
  (defun extrair-prefixo-letras (str / i prefixo ch)
    (setq i 1 prefixo "")
    (while (and (<= i (strlen str))
                (setq ch (substr str i 1))
                (wcmatch ch "[A-Za-z]"))
      (setq prefixo (strcat prefixo ch)
            i (1+ i)))
    prefixo
  )

  ;; === FUNÇÃO PARA LER CSV ===
  (defun ler-dados-csv ( / file line fields dados)
    (setq file (open ficheiroCsv "r"))
    (setq dados '())
    (if file
      (progn
        (read-line file) ; ignora cabeçalho
        (while (setq line (read-line file))
          (setq fields (split-string line ";"))
          (if (>= (length fields) 7)
            (progn
              (setq bore    (nth 0 fields))
              (setq x       (atof (nth 1 fields)))
              (setq y       (atof (nth 2 fields)))
              (setq z       (atof (nth 3 fields)))
              (setq formato (nth 5 fields))
              (setq cor     (nth 6 fields))
              (setq dados (cons (list bore x y z formato cor) dados))
            )
          )
        )
        (close file)
      )
    )
    dados
  )

  ;; === INSERIR OS BLOCOS ===
  (setq dados (ler-dados-csv))

  (setvar "ATTDIA" 0)
  (setvar "ATTREQ" 0)

  (princ "\n[2/3] A inserir blocos...")
  (foreach item dados
    (setq bore    (nth 0 item))
    (setq x       (nth 1 item))
    (setq y       (nth 2 item))
    (setq z       (nth 3 item))
    (setq formato (nth 4 item))
    (setq cor     (nth 5 item))
    (setq bloco   (strcat formato "_" cor))
    (setq pt      (list x y z))

    ;; === LAYER POR PREFIXO ===
    (setq prefixo (extrair-prefixo-letras bore))
    (setq layerNome (strcat "Log_" prefixo))

    ;; Cria a layer se necessário
    (if (not (tblsearch "LAYER" layerNome))
      (entmake (list (cons 0 "LAYER")
                     (cons 100 "AcDbSymbolTableRecord")
                     (cons 100 "AcDbLayerTableRecord")
                     (cons 2 layerNome)
                     (cons 70 0)
                     (cons 62 7))) ; cor branca
    )

    ;; Muda para a layer correta
    (setvar "CLAYER" layerNome)

    ;; Captura o entlast antes da inserção
    (setq antes (entlast))

    ;; Inserir bloco principal
    (command "_.-INSERT" bloco "_non" pt 1 1 0)

    ;; Captura o bloco inserido
    (setq depois (entlast))

    ;; Inserir bloco MIRA_cor
    (command "_.-INSERT" (strcat "MIRA_" cor) "_non" pt 0.025 0.025 0)

    ;; Atualizar atributo do bloco principal
    (if (and depois (tblsearch "BLOCK" bloco))
      (progn
        (setq blkData (entnext depois))
        (while (and blkData (/= (cdr (assoc 0 (entget blkData))) "SEQEND"))
          (if (= (cdr (assoc 0 (entget blkData))) "ATTRIB")
            (progn
              (setq attData (entget blkData))
              (entmod (subst (cons 1 bore) (assoc 1 attData) attData))
              (entupd blkData)
            )
          )
          (setq blkData (entnext blkData))
        )
      )
    )
  )

  (setvar "ATTDIA" 1)
  (setvar "ATTREQ" 1)
  (princ "\n[3/3] Concluído com sucesso. Blocos posicionados.")

  (princ)
)

