""##########################################################
""## Vim Email Client v0.2                                ##
""##                                                      ##
""## Created by David Elentok.                            ##
""##                                                      ##
""## This software is released under the GPL License.     ##
""##                                                      ##
""##########################################################

" ==========================================================
" Configuration
let g:mailboxdir = "d:\\data\\mailbox"
let g:sendscript = "d:\\dev\\python\\email\\send.py"
let g:fetchscript = "d:\\dev\\python\\email\\fetch.py"
let g:mailFrom = "Me <me@server1.com>"
let g:run_bat = "d:\\temp\\run.bat"



let g:addressbook = g:mailboxdir."\\addressbook.txt"



"" ==========================================================
"" dir = GetDir (path) 
""
""   this function returns the directory name from
""   a complete path ("c:\\temp\\file.txt" => "c:\\temp\\")
""
func! GetDir(path)
  let pos = strridx(a:path, "\\")
  if pos > 0
    return strpart(a:path, 0, pos+1)
  else
    return a:path
endfunc

"" ========================================================= 
"" Email_AttachFiles(path)
""
""   adds the file in 'path' to the email (adds the filename
""   with path after the '**** Attachments:' string in the
""   email message)
""   
""   multiple attachments can be added by giving the function
""   path = "c:\\temp\\*.txt" (for example), this can be done 
""   because the directory list is generated using the 'dir' 
""   command in windows, and the 'ls' command in linux)
""
func! Email_AttachFiles(path)
  let pos = search("#### Attachments:")

  if pos == 0
    normal Go#### Attachments:

  else
    normal o
  endif

  "" a:path comes with escaped spaces ('\ '), the dir command 
  "" doesn't work work with escaped characters, and it can't
  "" find the directory, so i'm removing changing every '\ '
  "" occurence to ' '.
  "" NOTE: '\\ ' != "\\ "
  ""  '' are without escaped characters, and "" are with escaped charaters
  let epath = substitute(a:path, '\\ ', " ","g")

  "" get the directory name with escaped ' ' and '\'
  let dir = escape(GetDir(epath),'\')

  "" get all files that match the regular expression (a:path)
  let list = system("dir /b \"".epath."\"")

  "" add the directory name in the beginning
  let list = substitute(list, "^", dir, "g")
  let list = substitute(list, "\n", "\n".dir, "g")

  "" add the list to the file
  exec "normal i=list"
  normal 0D
endfunc

"" =========================================================
"" command: Attach 
"" usage:
""     :Attach c:\temp\*.txt
command! -complete=file -nargs=+ Attach call Email_AttachFiles(<q-args>)

"" ==========================================================
"" Email_InsertNewHeader(header)
""
""   replaces the previous header fields(to, cc, bcc) in the 
""   .eml file with the new ones from Email_OpenAddressBook
""
func! Email_InsertNewHeader(header)
  normal gg
  if strpart(getline(1),0,3) == "To:"
    normal dd
  endif

  if strpart(getline(1),0,3) == "Cc:"
    normal dd
  endif

  if strpart(getline(1),0,4) == "Bcc:"
    normal dd
  endif

  exec "normal O".a:header
endfunc

"" =========================================================
"" Email_SetRecipients
""   this function is supposed to run after OpenAddressBook
""   from the addressbook buffer, what it does is create
""   a string that contains the new headers:
""      To: ...
""      Cc: ...
""      Bcc: ...
""   once the string is created the address book buffer is
""   closed.
""
func! Email_SetRecipients()
  %s/^[^\[].*\n//
  %s/^\n//
  normal GoEND

  let to = ""
  let cc = ""
  let bcc = ""

  let i = 1
  let line = getline(i)
  while line != "END"
    if strpart(line,0,4) == "[To]"
      if to != ""
        let to = to.", ".strpart(line, 4)
      else
        let to = strpart(line, 4)
      endif

    elseif strpart(line,0,4) == "[Cc]"
      if bcc != ""
        let cc = cc.", ".strpart(line, 4)
      else
        let cc = strpart(line, 4)
      endif

    elseif strpart(line,0,5) == "[Bcc]"
      if bcc != ""
        let bcc = bcc.", ".strpart(line, 5)
      else
        let bcc = strpart(line, 5)
      endif

    endif

    let i = i+1
    let line = getline(i)
  endwhile

  bd!
  let header = "To: ".to
  if cc != ""
    let header = header."\nCc: ".cc
  endif
  
  if bcc != ""
    let header = header."\nBcc: ".bcc
  endif

  call Email_InsertNewHeader(header)
endfunc



"" =========================================================
"" Email_GetRecipients
""
""   this function is to be run inside an .eml file, it returns
""   the 'To', 'Cc' and 'Bcc' fields for usage with OpenAddressBook
func! Email_GetRecipients()
  let fields = ""
  if strpart(getline(1), 0, 3) == "To:"
    let fields = getline(1)
  endif

  if strpart(getline(2), 0, 3) == "Cc:"
    let fields = fields."\n".getline(2)
  elseif strpart(getline(2), 0, 4) == "Bcc:"
    let fields = fields."\n".getline(2)
  endif

  if strpart(getline(3), 0, 4) == "Bcc:"
    let fields = fields."\n".getline(3)
  endif

  return fields
endfunc

"" =========================================================
"" Email_SelectRecipients
""
""   this function is to be run inside the address book buffer,
""   it selects all of the recipients inside the argument 
""   'fields' (To: .. Cc: .. Bcc: ..)
""
func! Email_SelectRecipients(fields)
  if a:fields == ""
    return

  else
    let str = a:fields."\n"
    let pos = stridx(str,"\n")
    while pos != -1
      let line = strpart(str,0,pos)
      let str = strpart(str,pos+1)

      let prefix=""
      if strpart(line,0,3) == "To:"
        let line = strpart(line,3).","
        let prefix="[To]"
      elseif strpart(line,0,3) == "Cc:"
        let line = strpart(line,3).","
        let prefix="[Cc]"
      elseif strpart(line,0,4) == "Bcc:"
        let line = strpart(line,4).","
        let prefix="[Bcc]"
      endif

      if prefix != ""
        let pos2 = stridx(line,",")
        while pos2 != -1
          let email = strpart(line, 1, pos2-1)
          let line = strpart(line, pos2+1)
"           echo prefix." -> [".email."]"
          let linenumber = search(email)
          exec "normal 0i".prefix
          let pos2 = stridx(line,",")
        endwhile
      endif

      let pos = stridx(str,"\n")
    endwhile
  endif
endfunc

"" =========================================================
"" Email_OpenAddressBook
""   opens an address book buffer, allowing the user to select
""   recipients
func! Email_OpenAddressBook()
  let email_mode = 0
  if expand("%:e") == "eml"
    let email_mode = 1
    let fields = Email_GetRecipients()
  endif
    
  new AddressBook
  exec "read ".g:addressbook
  normal ggOHelp: t=[To], T=[Cc], b=[Bcc], Enter=save, q=close

  if email_mode
    call Email_SelectRecipients(fields)
  endif

  map <buffer> t 0i[To]<esc>
  map <buffer> T 0i[Cc]<esc>
  map <buffer> b 0i[Bcc]<esc>
  map <buffer> q :bd!<cr>
  map <buffer> <cr> :call Email_SetRecipients()<cr>
  syn region headerTo start="^\[To" end="$"
  syn region headerCc start="^\[Cc" end="$"
  syn region headerBcc start="^\[Bcc" end="$"
  syn region email start="<" end=">" end="$"
  syn region help start="^Help:" end="$"
  hi headerTo guifg=green
  hi headerCc guifg=lightblue
  hi headerBcc guifg=orange
  hi email guifg=yellow
  hi help guifg=magenta
endfunc

" ==========================================================
" Email_Compose
func! Email_Compose()
  let subject = input("Enter Subject: ")
  exec "new ".g:mailboxdir."\\temp.eml"
  exec "normal ggVGxiFrom: ".g:mailFrom."\nSubject: ".subject
  call Email_OpenAddressBook()
endfunc

" ==========================================================
" Email_Send
func! Email_Send()
  let yesno = input("Are you sure you want to send [y/n]? ")
  if yesno == "y"
    exec ":!".g:sendscript." %"
  endif
endfunc

" ==========================================================
" Email_FoldingLevel
func! Email_FoldingLevel(n)
  let thisLine = getline(a:n)
  if strpart(thisLine,0,10) == "##########"
    return '>1'
  endif

  return '='
endfunc

" ==========================================================
" Email_FoldingText
func! Email_FoldingText()
  if v:foldlevel==1
    return getline(v:foldstart+1)
  endif
endfunc

" ==========================================================
" Email_MailboxMode
func! Email_MailboxMode()
"if &foldexpr != "Email_FoldingLevel(v:lnum)"
  if &l:foldtext != "Email_FoldingText()"
    setlocal foldexpr=Email_FoldingLevel(v:lnum)
    setlocal foldtext=Email_FoldingText()
    setlocal foldmethod=expr foldlevel=0 foldcolumn=2
    setlocal foldenable
  endif
  setlocal syntax=mail
  map <cr>          :call Email_JumpToLink()<cr>
  map <backspace>   :foldclose<cr>
  map <c-backspace> :setlocal foldlevel=0<cr>
  syn region separatorLine start="#####" end="$"
  syn region attachments start="#### Attachments" end="$"
  hi separatorLine guifg=green guibg=blue
  hi attachments guifg=black guibg=gray
endfunc

" ==========================================================
" Email_ComposeMode
func! Email_ComposeMode()
  setf mail
  setlocal nocindent autoindent
  setlocal textwidth=70
  syn region attachments start="#### Attachments" end="$"
  hi attachments guifg=black guibg=gray
endfunc

" ==========================================================
" Email_JumpToLink
func! Email_JumpToLink()
  if !filereadable(g:run_bat)
    exec '!echo start "title" "\%*" > '.g:run_bat
  endif

  normal 0v$"3y

  let lnk=@3
  "TODO
  let lnk=substitute(lnk, "%attachments%", g:mailboxdir."\\attach", "g")

  echo "link: ".lnk
  if lnk =~ '.*\.py' || lnk =~ '.*\.txt' || lnk =~ '.*\.vim' || lnk =~ '.*\.asm' || lnk =~ '^[^.]*$'
    exec "e ".lnk
  else
    exec "silent !start ".g:run_bat." ".lnk
  endif

endfunc

" ==========================================================
" Init
if !isdirectory(g:mailboxdir)
  exec "!mkdir ".g:mailboxdir
endif

" ==========================================================
" Auto Commands
au BufNewFile,BufRead *.mbx call Email_MailboxMode()
au BufNewFile,BufRead *.eml call Email_ComposeMode()



" ==========================================================
" Keymapping
map za :call Email_OpenAddressBook()<cr>
map zs :call Email_Send()<cr>
map zc :call Email_Compose()<cr>
map zf :call search("#### Attachments:")<cr>

" open last sent email:
map zC :exec "e ".g:mailboxdir."\\temp.eml"<cr>
