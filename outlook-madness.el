;; never do html->plaintext conversion (will probably go wrong eventually...)
;; using w3m for conversion will make horizontal rules utf8 chars which behave
;; poorly with outlook reply style
(setq mu4e-view-html-plaintext-ratio-heuristic most-positive-fixnum)

(defun my/mu4e-pre-compose-detect-reply-fields ()
  "When composing, detect the 'To' and 'Cc' fields and store
   semicolon-separated list of recipient names in
   my/mu4e-msg-reply-field (for outlook reply style)."
  (when (> (length mu4e-compose-parent-message) 0)
    (setq my/mu4e-parent-message mu4e-compose-parent-message)
    ))
(add-hook 'mu4e-compose-pre-hook 'my/mu4e-pre-compose-detect-reply-fields)

;; save id of message being replied to in local my/mu4e-parent-message-id variable
(defun my/mu4e-post-compose-set-parent-message-id ()
  (when (> (length mu4e-compose-parent-message) 0)
    (make-local-variable 'my/mu4e-parent-message-id)
    (setq my/mu4e-parent-message-id (mu4e-message-field my/mu4e-parent-message :message-id))))
(add-hook 'mu4e-compose-mode-hook 'my/mu4e-post-compose-set-parent-message-id)

;; use external python script to format outlook-style reply
(defun my/mu4e-outlook-madness ()
  (interactive)
  (save-excursion
    (message-goto-body)
    (insert my/mu4e-parent-message-id)
    (newline)
    (message-goto-body)
    (shell-command-on-region (point) (point-max) "mimedown.py" nil t)
    (browse-url "file:///dev/shm/mimedown-madness.html")
    ))
(add-hook 'mu4e-compose-mode-hook (lambda () (local-set-key (kbd "C-c m") 'my/mu4e-outlook-madness)))

;;; opening browser to show mimedown preview - use qutebrowser to do so, here:
;; define function for browsing urls with qutebrowser
(defun browse-url-qutebrowser (url &optional _new-window)
  ""
  (interactive (browse-url-interactive-arg "URL: "))
  (setq url (browse-url-encode-url url))
  (let* ((process-environment (browse-url-process-environment)))
    (apply 'start-process
	   (concat "qutebrowser " url) nil
	   "qutebrowser"
	   (append
	    nil
	    (list url)))))

;; use qutebrowser to browse urls
(setq browse-url-browser-function 'browse-url-qutebrowser)

;; confirmation before send
(add-hook 'message-send-hook
  (lambda ()
    (unless (yes-or-no-p "Send mail?")
      (signal 'quit nil))))
