;; ver. 0.0.1

(vl-load-com)

(defun c:getla (/ gl)

               ;|
(gl "E:\\Repair" "C:\\Users\\a.kulik\\Documents\\Drawing1.txt")
(gl nil nil)
|;

  (defun gl (path                      file
             /                         lst
             *kpblc-acad*              *kpblc-adoc*
             dwg                       file
             mask                      odbx
             path                      vl-browsefiles-in-directory-nested
             vl-browsefolder           _kpblc-dir-path-and-splash
             _kpblc-odbx-open          _kpblc-odbx-close
             _kpblc-odbx               _kpblc-eval-value-round
             _kpblc-conv-value-to-string
             _kpblc-conv-value-to-int  _kpblc-acad-version
             _kpblc-conv-vla-to-list   _kpblc-conv-ent-to-vla
             _kpblc-conv-ent-to-ename  _kpblc-is-file-read-only
             )

    (defun _kpblc-conv-vla-to-list (value / res)
                                   ;|
*    Преобразовывает vlax-variant или vlax-safearray в список.
|;
      (cond
        ((listp value)
         (mapcar (function _kpblc-conv-vla-to-list) value)
         )
        ((= (type value) 'variant)
         (_kpblc-conv-vla-to-list (vlax-variant-value value))
         )
        ((= (type value) 'safearray)
         (if (>= (vlax-safearray-get-u-bound value 1) 0)
           (_kpblc-conv-vla-to-list (vlax-safearray->list value))
           ) ;_ end of if
         )
        ((and (member (type value) (list 'ename 'str 'vla-object))
              (= (type (_kpblc-conv-ent-to-vla value)) 'vla-object)
              (vlax-property-available-p (_kpblc-conv-ent-to-vla value) 'count)
              ) ;_ end of and
         (vlax-for sub (_kpblc-conv-ent-to-vla value)
           (setq res (cons sub res))
           ) ;_ end of vlax-for
         )
        (t value)
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-acad-version ()
                               ;|
*    Определение номера сборки AutoCAD
*    Возвращаемое значение: Число двойной точности. Для AutoCAD 2005 вернет 16.1, для 2006 - 16.2 и т.д.
Примеры вызова:
(_kpblc-acad-version)
|;
      (atof (getvar "acadver"))
      ) ;_ end of defun

    (defun _kpblc-conv-value-to-int (value /)
                                    ;|
*    конвертация значения в целое. Для VLA-объектов возвращается nil.
*    Точечные списки не обрабатываются.
|;
      (cond
        ((not value) 0)
        ((equal value t) 1)
        (t (atoi (_kpblc-conv-value-to-string value)))
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-conv-value-to-string (value /)
                                       ;|
*    конвертация значения в строку.
|;
      (cond
        ((= (type value) 'str) value)
        ((= (type value) 'int) (itoa value))
        ((and (= (type value) 'real) (equal value (_kpblc-eval-value-round value 1.) 1e-6))
         (itoa (fix value))
         )
        ((= (type value) 'real) (rtos value 2 14))
        ((not value) "")
        (t (vl-princ-to-string value))
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-eval-value-round (value to)
                                   ;|
;; http://forum.dwg.ru/showthread.php?p=301275
*    Выполняет округление числа до указанной точности
*    Примеры вызова:
(_kpblc-eval-value-round 16.365 0.01) ; 16.37
|;
      (if (zerop to)
        value
        (* (atoi (rtos (/ (float value) to) 2 0)) to)
        ) ;_ end of if
      ) ;_ end of defun

    (defun _kpblc-odbx (/)
                       ;|
*    функция возвращает интерфейс IAxDbDocument (для работы с файлами DWG без
* их открытия). Если интерфейс не поддерживается, возвращает nil. Проверено
* на ACAD 2002, 2004, 2005, 2006, 2007
*    Автор - Fatty aka Олег jr. Моего только адаптация под общую систему и
* переименование
*    Параметры вызова:
*  нет
*    Примеры вызова:
(_kpblc-odbx)
|;
      (cond
        ((< (_kpblc-acad-version) 15.06)
         (alert
           "ObjectDBX method not applicable\nin this AutoCAD version"
           ) ;_ end of KPBLC-MSG-ALERT
         nil
         )
        ((= (fix (_kpblc-acad-version)) 15)
         (if (not (vl-registry-read
                    "HKEY_CLASSES_ROOT\\ObjectDBX.AxDbDocument\\CLSID"
                    ) ;_ end of vl-registry-read
                  ) ;_ end of not
           (startapp "regsvr32.exe"
                     (strcat "/s \"" (findfile "axdb15.dll") "\"")
                     ) ;_ end of startapp
           ) ;_ end of if
         (vla-getinterfaceobject
           (vlax-get-acad-object)
           "ObjectDBX.AxDbDocument"
           ) ;_ end of vla-getinterfaceobject
         )
        (t
         (vla-getinterfaceobject
           (vlax-get-acad-object)
           (strcat "ObjectDBX.AxDbDocument."
                   (_kpblc-conv-value-to-string
                     (_kpblc-conv-value-to-int (_kpblc-acad-version))
                     ) ;_ end of _kpblc-conv-value-to-string
                   ) ;_ end of strcat
           ) ;_ end of vla-getinterfaceobject
         )
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-odbx-close (conn)
                             ;|
*    Закрытие файла, открытого ранее через _kpblc-odbx-*. С попыткой сохранения
*    Параметры вызова:
*  conn  соединение с ObjectDBX, созданное ранее через (_kpblc-odbx)
*    либо список:
      '(("conn" . <ObjectDBXConnection>)  ; то же самое
  ("save" . t)      ; сохранять или нет изменения
  ("file" . "c:\\temp\\tmp.dwg")  ; имя, под которым сохранять. nil ->
    ; использовать текущее
|;
      (if (and (= (type conn) 'list)
               (cdr (assoc "save" conn))
               ) ;_ end of and
        (progn
          (vlax-invoke
            (cdr (assoc "conn" conn))
            'saveas
            (cond
              ((cdr (assoc "file" conn))
               (strcat
                 (_kpblc-dir-path-and-splash (vl-filename-directory file))
                 (_kpblc-string-ext (vl-filename-base file) "dwg")
                 ) ;_ end of strcat
               )
              (t (vla-get-name (cdr (assoc "conn" conn))))
              ) ;_ end of cond
            ) ;_ end of vlax-invoke
          ) ;_ end of progn
        ) ;_ end of if
      (vl-catch-all-apply
        '(lambda ()
           (vlax-release-object
             (if (= (type conn) 'list)
               (cdr (assoc "conn" conn))
               conn
               ) ;_ end of if
             ) ;_ end of vlax-release-object
           ) ;_ end of lambda
        ) ;_ end of vl-catch-all-apply
      (setq conn nil)
      ) ;_ end of defun

    (defun _kpblc-odbx-open (file odbx / res obj tmp_file)
                            ;|
*    Открытие любого файла, даже в режиме "ReadOnly"
*    Параметры вызова:
  file  полное имя открываемого файла. Только строка, контроля не
    выполняется
  odbx  ObjectDBX-интерфейс, созданный (_kpblc-odbx).
*    Возвращает список вида:
  '(("obj" . <vla-указатель на гарантированно открытый документ>)
    ("close" . t | nil)  ; допускается ли закрытие файла
    ("save" . t | nil)  ; допускается ли сохранение файла
    ("write" . t | nil)  ; допускается ли запись в файл
    ("name" . <строка имени файла>))
|;
      (cond
        ((not file)
         (setq res (list (cons "obj" *kpblc-adoc*)
                         (cons "write" t)
                         (cons "name" (vla-get-fullname *kpblc-adoc*))
                         ) ;_ end of list
               ) ;_ end of setq
         )
        ((member
           (strcase file)
           (mapcar (function (lambda (x) (strcase (vla-get-fullname x))))
                   (_kpblc-conv-vla-to-list
                     (vla-get-documents *kpblc-acad*)
                     ) ;_ end of _kpblc-conv-vla-to-list
                   ) ;_ end of mapcar
           ) ;_ end of member
         (setq
           res (list
                 (cons "obj"
                       (car (vl-remove-if-not
                              '(lambda (x)
                                 (= (strcase (vla-get-fullname x)) (strcase file))
                                 ) ;_ end of lambda
                              (_kpblc-conv-vla-to-list
                                (vla-get-documents *kpblc-acad*)
                                ) ;_ end of _kpblc-conv-vla-to-list
                              ) ;_ end of vl-remove-if-not
                            ) ;_ end of car
                       ) ;_ end of cons
                 (cons "write" t)
                 (cons "save" t)
                 (cons "name" file)
                 ) ;_ end of list
           ) ;_ end of setq
         )
        ((and (findfile file)
              (_kpblc-is-file-read-only file)
              ) ;_ end of and
         (vl-file-copy
           file
           (setq tmp_file
                  (strcat
                    (vl-filename-mktemp
                      (strcat (vl-filename-base file)
                              (vl-filename-extension file)
                              ) ;_ end of strcat
                      ) ;_ end of vl-filename-mktemp
                    ) ;_ end of strcat
                 ) ;_ end of setq
           ) ;_ end of vl-file-copy
         (vla-open odbx tmp_file)
         (setq res (list (cons "obj" odbx)
                         (cons "close" t)
                         (cons "save" nil)
                         (cons "write" nil)
                         (cons "name" file)
                         ) ;_ end of list
               ) ;_ end of setq
         )
        ((and (findfile file)
              (not (_kpblc-is-file-read-only file))
              ) ;_ end of and
         (vla-open odbx file)
         (setq res (list (cons "obj" odbx)
                         (cons "close" t)
                         (cons "save" t)
                         (cons "write" t)
                         (cons "name" file)
                         ) ;_ end of list
               ) ;_ end of setq
         )
        ) ;_ end of cond
      res
      ) ;_ end of defun

    (defun _kpblc-dir-path-and-splash (path)
                                      ;|
*    Возвращает путь со слешем в конце
*    Параметры вызова:
*  path  - обрабатываемый путь
*    Примеры вызова:
(_kpblc-dir-path-and-splash "c:\\kpblc-cad")  ; "c:\\kpblc-cad\\"
|;
      (strcat (vl-string-right-trim "\\" path) "\\")
      ) ;_ end of defun

    (defun vl-browsefolder (caption / shlobj folder fldobj outval)
                           ;|
http://www.autocad.ru/cgi-bin/f1/board.cgi?t=21054YY    
*    Без отображения файлов
*    Параметры вызова:
	caption		показываемый заголовок (пояснение) окна
(setq Folder (vlax-invoke-method ShlObj 'BrowseForFolder 0 "" 16384))
|;
      (setq shlobj (vla-getinterfaceobject
                     (vlax-get-acad-object)
                     "Shell.Application"
                     ) ;_ end of vla-getInterfaceObject
            folder (vlax-invoke-method shlobj 'browseforfolder 0 caption 0)
            ) ;_ end of setq
      (vlax-release-object shlobj)
      (if folder
        (progn (setq fldobj (vlax-get-property folder 'self)
                     outval (vlax-get-property fldobj 'path)
                     ) ;_ end of setq
               (vlax-release-object folder)
               (vlax-release-object fldobj)
               ) ;_ end of progn
        ) ;_ end of if
      outval
      ) ;_ end of defun

    (defun vl-browsefiles-in-directory-nested (path mask)
                                              ;|
*    Функция возвращает список файлов указанной маски, находящихся в
* заданном каталоге
*    Параметры вызова:
  path  путь к корневому каталогу. nil недопустим
  mask  маска имени файла. nil или список недопустим
*    Примеры вызова:
(vl-browsefiles-in-directory-nested "c:\\documents" "*.dwg")
|;
      (apply
        (function append)
        (cons
          (if (vl-directory-files path mask)
            (mapcar
              (function (lambda (x)
                          (strcat (vl-string-right-trim "\\" path) "\\" x)
                          ) ;_ end of lambda
                        ) ;_ end of function
              (vl-directory-files path mask)
              ) ;_ end of mapcar
            ) ;_ if
          (mapcar (function
                    (lambda (x)
                      (vl-browsefiles-in-directory-nested
                        (strcat (vl-string-right-trim "\\" path) "\\" x)
                        mask
                        ) ;_ end of vl-browsefiles-in-directory-nested
                      ) ;_ end of lambda
                    ) ;_ end of function
                  (vl-remove ".."
                             (vl-remove "." (vl-directory-files path nil -1))
                             ) ;_ end of vl-remove
                  ) ;_ mapcar
          ) ;_ cons
        ) ;_ end of apply
      ) ;_ end of defun

    (defun _kpblc-conv-ent-to-vla (ent_value / res)
                                  ;|
*    Функция преобразования полученного значения в vla-указатель.
*    Параметры вызова:
*  ent_value  значение, которое надо преобразовать в указатель. Может
*      быть именем примитива, vla-указателем или просто
*      списком.
*      Если не принадлежит ни одному из указанных типов,
*      возвращается nil
*    Примеры вызова:
(_kpblc-conv-ent-to-vla (entlast))
(_kpblc-conv-ent-to-vla (vlax-ename->vla-object (entlast)))
|;
      (cond
        ((= (type ent_value) 'vla-object) ent_value)
        ((= (type ent_value) 'ename) (vlax-ename->vla-object ent_value))
        ((setq res (_kpblc-conv-ent-to-ename ent_value))
         (vlax-ename->vla-object res)
         )
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-conv-ent-to-ename (ent_value /)
                                    ;|
*    Функция преобразования полученного значения в ename
*    Параметры вызова:
*  ent_value  значение, которое надо преобразовать в примитив. Может
*      быть именем примитива, vla-указателем или просто
*      списком.
*      Если не принадлежит ни одному из указанных типов,
*      возвращается nil
*    Примеры вызова:
(_kpblc-conv-ent-to-ename (entlast))
(_kpblc-conv-ent-to-ename (vlax-ename->vla-object (entlast)))
|;
      ;; "_kpblc-conv-ent-to-ename")
      (cond
        ((= (type ent_value) 'vla-object)
         (vlax-vla-object->ename ent_value)
         )
        ((= (type ent_value) 'ename) ent_value)
          ;((= (type ent_value) 'str) (handent ent_value))
        ((= (type ent_value) 'list) (cdr (assoc -1 ent_value)))
        (t nil)
        ) ;_ end of cond
      ) ;_ end of defun

    (defun _kpblc-is-file-read-only (file-name / file_hangle res)
                                    ;|
*    Проверяет, является ли файл "read-only". Возвращает t, если да. Проверки
* наличия файла не выполняется.
*    Параметры вызова:
*  file-name  полное имя файла, с путем.
(_kpblc-is-file-read-only "Z:\\КТО transit\\Разное\\Устройство молниезащиты.dwg")
|;
      (and file-name
           (findfile file-name)
           (or (not (vl-file-systime file-name))
               ((lambda (/ svr obj res)
                  (setq svr (vlax-get-or-create-object "Scripting.FileSystemObject")
                        obj (vlax-invoke-method svr 'getfile file-name)
                        res (vlax-get-property obj 'attributes)
                        ) ;_ end of setq
                  (vlax-release-object obj)
                  (vlax-release-object svr)
                  (setq obj nil
                        svr nil
                        ) ;_ end of setq
                  (/= (* 2 (/ res 2)) res)
                  ) ;_ end of lambda
                )
               ) ;_ end of or
           ) ;_ end of and
      ) ;_ end of defun

    (if (and (or path
                 (and (setq path (vl-browsefolder "Выберите обрабатываемый каталог"))
                      (/= path "")
                      ) ;_ end of and
                 ) ;_ end of or
             (or file (setq file (getfiled "Куда сохранять результаты" "" "txt" 1)))
             ) ;_ end of and
      (progn
        (setq *kpblc-adoc* (vla-get-activedocument (setq *kpblc-acad* (vlax-get-acad-object))))
        (foreach item (vl-browsefiles-in-directory-nested path "*.dwg")
          (setq dwg (_kpblc-odbx-open item (setq odbx (_kpblc-odbx))))
          (vlax-for layer (vla-get-layers (cdr (assoc "obj" dwg)))
            (if (not (member (strcase (vla-get-name layer)) lst))
              (setq lst (cons (strcase (vla-get-name layer)) lst))
              ) ;_ end of if
            ) ;_ end of vlax-for
          (_kpblc-odbx-close odbx)
          ) ;_ end of foreach
        (setq handle (open file "w"))
        (foreach str (vl-sort lst '<) (write-line str handle))
        (close handle)
        (command "_.shell" (strcat "\"" file "\""))
        ) ;_ end of progn
      ) ;_ end of if
    ) ;_ end of defun

  (gl nil nil)
  (princ)
  ) ;_ end of defun
