;; WARNING: This is a terrible hack written in an attempt to support the
;; terrible hacks implemented by Outlook. The author takes no responsibility and
;; users are encouraged to closely inspect the generated message before sending!

;; Not much is going on in lisp: We need a reference to the parent message (the
;; message being replied to) so we can forward its ID to mimelook.py. mu4e
;; exposes the message being replied to in the variable
;; `mu4e-compose-parent-message' in the pre-compose hook, so we grab it, store
;; it in `mimelook/mu4e-parent-message', and then store it in a local variable
;; in the `mu4e-compose-mode-hook'.

;; Before sending, we call `mimelook/mu4e-outlook-madness'. This will take the
;; current body text and parent message ID from the buffer and send it to an
;; external Python script (mimelook.py). This script then does several things:

;; 1) It generates an HTML version of the body text (assuming it is Markdown
;; syntax, with the rationale that "the single biggest source of inspiration for
;; Markdown’s syntax is the format of plain text email" [1]).

;; 2) It finds the parent message using mu and the message id and parses it
;; using the Python mailparser module.

;; 3) It fetches all inline attachments (usually images) of the parent message,
;; which is commonly used in Outlook users' signatures, and stores them in a
;; directory in /dev/shm/.

;; 4) It grabs the HTML body text of the parent message

;; Finally, a multipart message is generated with the structure:

;; <#multipart type=alternative>
;;  <<untouched plaintext that was in the buffer before calling `mimelook/mu4e-outlook-madness'>>
;; <#part type=text/html>
;;  <<everything in the parent message HTML upto and including the <body> tag>>
;;  <<the plaintext from our buffer, converted to HTML via Markdown>>
;;  <<an Outlook-style message citation block with "From: x, To: y, Subject: z, ...">>
;;  <<everything in the parent message HTML from (but excluding) the <body> tag>>
;; <#/multipart>
;;  <<a <#part ...><#/part> for each inline attachment>>

;; This message overwrites the content of the buffer, which can then be sent as
;; usual (usually with C-c C-c).

;; The result is a message that has a nice, sane plaintext version as we wrote
;; it in mu4e. It also has an insane HTML version that renders nicely for
;; Outlook users and contains the entire mail thread being replied to below the
;; message being sent, as Outlook users usually expect.

;; For debugging, the HTML part is written to /dev/shm/mimelook-madness.html for
;; inspection. This file will automatically be opened in a browser if
;; `mimelook/mu4e-outlook-madness' is called with a prefix argument. Do note
;; that inline attachment images are NOT rendered in this preview!

;; If an Outlook user sends us a plaintext email, don't attempt to produce an
;; HTML reply with `mimelook/mu4e-outlook-madness'. Instead refer to
;; `mimelook/enable-outlook-reply-style', which will set
;; `message-citation-line-function' to an Outlook-style (but plaintext) citation
;; line (will also change yank prefixes). These are bound to C-c C-S-o and C-c
;; C-o.

;; [1]: https://daringfireball.net/projects/markdown/

(defun mimelook/mu4e-pre-compose-set-parent-message ()
  "When composing, detect the message being replied to and store
it in the global (and unsafe to use) variable
`mimelook/mu4e-parent-message'."
  (when (> (length mu4e-compose-parent-message) 0)
    (setq mimelook/mu4e-parent-message mu4e-compose-parent-message)))
(add-hook 'mu4e-compose-pre-hook 'mimelook/mu4e-pre-compose-set-parent-message)

(defun mimelook/mu4e-post-compose-set-parent-message ()
  "When composing, use `mimelook/mu4e-parent-message' to store the parent
message in the local variable `mimelook/mu4e-local-parent-message'."
  (message "mimelook/mu4e-post-compose-set-parent-message")
  (when (> (length mu4e-compose-parent-message) 0)
    (make-local-variable 'mimelook/mu4e-local-parent-message)
    (setq mimelook/mu4e-local-parent-message mu4e-compose-parent-message)))
(add-hook 'mu4e-compose-mode-hook 'mimelook/mu4e-post-compose-set-parent-message)

(defun mimelook/mu4e-outlook-madness (x)
  "Convert message in current buffer to a multipart message where
the HTML version loosely conforms to the style used by MS
outlook. This depends on the external script mimelook.py that
has to be in PATH. With prefix argument, open browser for
preview.

Not compatible with `mimelook/message-insert-outlook-citation-line'."
  (interactive "P")
  (save-excursion
    (message-goto-body)
    (insert (mu4e-message-field mimelook/mu4e-local-parent-message :message-id))
    (newline)
    (message-goto-body)
    (shell-command-on-region (point) (point-max) "mimelook.py" nil t)
    (when x
      (message "Opening HTML preview - inline images are NOT rendered!")
      (browse-url "file:///dev/shm/mimelook-madness.html"))))
;; when in mu4e-compose mode, bind mimelook/mu4e-outlook-madness to C-c m
(add-hook 'mu4e-compose-mode-hook (lambda () (local-set-key (kbd "C-c m") 'mimelook/mu4e-outlook-madness)))

;; confirmation before send in case we forget calling mimelook/mu4e-outlook-madness
(add-hook 'message-send-hook
          (lambda ()
            (unless (yes-or-no-p "Send mail?")
              (signal 'quit nil))))

(defun mimelook/message-insert-outlook-citation-line ()
  "Based off `message-insert-citation-line'. Inserts outlook (web)-style replies.

Not compatible with `mimelook/mu4e-outlook-madness'"
  (when message-reply-headers
    (newline)
    (insert "________________________________")
    (newline)
    (insert "From: " (replace-regexp-in-string " <.*>" "" (mail-header-from message-reply-headers)))
    (newline)
    (insert "Sent: "
	    (let ((tstr (parse-time-string
			 (substring
			  (replace-regexp-in-string "T" " " (mail-header-date message-reply-headers))
			  0 -5))))
	      (format-time-string "%e %b %G %H:%M:%S" (apply 'encode-time tstr))))
    (newline)
    (setq value "")
    (insert "To: " (substring (concat (dolist (elt (mu4e-message-field mimelook/mu4e-parent-message :to) value)
					(setq value (concat value (car elt) "; "))))
			      0 -2))
    (when (> (length (mu4e-message-field mimelook/mu4e-parent-message :cc)) 0)
      (newline)
      (setq value "")
      (insert "Cc: " (substring (concat (dolist (elt (mu4e-message-field mimelook/mu4e-parent-message :cc) value)
					  (setq value (concat value (car elt) "; "))))
				0 -2)))
    (newline)
    (insert "Subject: " (mail-header-subject message-reply-headers))
    (newline)(newline)(newline)))

(defun mimelook/enable-outlook-reply-style ()
  "Enable Outlook-style replies. The inverse of
`mimelook/disable-outlook-reply-style'."
  (interactive)
  (setq message-citation-line-function 'mimelook/message-insert-outlook-citation-line)
  (setq message-yank-prefix "")
  (setq message-yank-empty-prefix "")
  (setq message-yank-cited-prefix "")
  (message "Enabled Outlook reply style"))

(defun mimelook/disable-outlook-reply-style ()
  "Disable Outlook-style replies. The inverse of
`mimelook/enable-outlook-reply-style'."
  (interactive)
  (setq message-citation-line-function 'message-insert-citation-line)
  (setq message-yank-prefix "> ")
  (setq message-yank-empty-prefix ">")
  (setq message-yank-cited-prefix ">")
  (message "Disabled Outlook reply style"))

;; hotkeys for enabling/disabling outlook reply styles when in mu4e-view-mode
(add-hook 'mu4e-view-mode-hook (lambda () (local-set-key (kbd "C-c C-S-o") 'mimelook/enable-outlook-reply-style)))
(add-hook 'mu4e-view-mode-hook (lambda () (local-set-key (kbd "C-c C-o") 'mimelook/disable-outlook-reply-style)))

;; initially disable outlook plaintext reply style
(mimelook/disable-outlook-reply-style)
