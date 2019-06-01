;; WARNING: This is a terrible hack written in an attempt to support the
;; terrible hacks implemented by Outlook. The author takes no responsibility and
;; users are encouraged to closely inspect the generated message before sending!

;; Not much is going on in lisp: We need a reference to the parent message (the
;; message being replied to) so we can forward its ID to mimelook.py. mu4e
;; exposes the message being replied to in the variable
;; mu4e-compose-parent-message in the pre-compose hook, so we grab it, store it
;; in my/mu4e-parent-message, and then store its ID in a local variable in the
;; mu4e-compose-mode-hook.

;; Before sending, we call my/mu4e-outlook-madness. This will take the current
;; body text and parent message ID from the buffer and send it to an external
;; Python script (mimelook.py). This script then does several things:

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
;;  <<untouched plaintext that was in the buffer before calling my/mu4e-outlook-madness>>
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
;; my/mu4e-outlook-madness is called with a prefix argument. Do note that inline
;; attachment images are NOT rendered in this preview!

;; [1]: https://daringfireball.net/projects/markdown/

(defun my/mu4e-pre-compose-set-parent-message ()
  "When composing, detect the message being replied to and store
   it in the global (and thus unsafe to use) variable
   my/mu4e-parent-message."
  (when (> (length mu4e-compose-parent-message) 0)
    (setq my/mu4e-parent-message mu4e-compose-parent-message)))
(add-hook 'mu4e-compose-pre-hook 'my/mu4e-pre-compose-set-parent-message)

(defun my/mu4e-post-compose-set-parent-message-id ()
  "When composing, use my/mu4e-parent-message to store the id of
   the parent message in the local variable
   my/mu4e-parent-message-id."
  (when (> (length mu4e-compose-parent-message) 0)
    (make-local-variable 'my/mu4e-parent-message-id)
    (setq my/mu4e-parent-message-id (mu4e-message-field my/mu4e-parent-message :message-id))))
(add-hook 'mu4e-compose-mode-hook 'my/mu4e-post-compose-set-parent-message-id)

(defun my/mu4e-outlook-madness (x)
  "Convert message in current buffer to a multipart message where
   the HTML version loosely conforms to the style used by MS
   outlook. This depends on the external script mimelook.py that has
   to be in PATH. With prefix argument, open browser for preview."
  (interactive "P")
  (save-excursion
    (message-goto-body)
    (insert my/mu4e-parent-message-id)
    (newline)
    (message-goto-body)
    (shell-command-on-region (point) (point-max) "mimelook.py" nil t)
    (when x
      (message "Opening HTML preview - inline images are NOT rendered!")
      (browse-url "file:///dev/shm/mimelook-madness.html"))))
;; when in mu4e-compose mode, bind my/mu4e-outlook-madness to C-c m
(add-hook 'mu4e-compose-mode-hook (lambda () (local-set-key (kbd "C-c m") 'my/mu4e-outlook-madness)))

;; confirmation before send in case we forget calling my/mu4e-outlook-madness
(add-hook 'message-send-hook
  (lambda ()
    (unless (yes-or-no-p "Send mail?")
      (signal 'quit nil))))
