" vim:foldmethod=marker:fen:
scriptencoding utf-8

" Saving 'cpoptions' {{{
let s:save_cpo = &cpo
set cpo&vim
" }}}


function! excel#load()
    " dummy function to load this script.
endfunction

function! excel#edit(file, bang)
    let book = s:Excel.Book.new(a:file)
    call s:open_buffer(book, a:bang)
endfunction

function! s:open_buffer(book, bang)
    " Behave like :edit command.
    if &modified && !a:bang
        call s:error('No write since last change (add ! to overwrite)')
        return
    endif

    enew!
    silent file `=a:book.get_filename()`
    setlocal buftype=nofile bufhidden=unload noswapfile nobuflisted
    setfiletype excel

    " Load book content.
    try
        call setline(1, a:book.get_lines())
    catch /^excel:/
        call s:error('Cannot read excel content.')
        call s:error(v:exception . ' ' . v:throwpoint)
    catch
        call s:error('Cannot set excel content to buffer.')
        call s:error(v:exception . ' ' . v:throwpoint)
    endtry
endfunction





let s:WSH_SCRIPT = globpath(&rtp, 'macros/excel-com.js')
let s:IS_MSWIN = has('win16') || has('win32') || has('win64') || has('win95')


let s:Excel = {}
let s:Excel.Book = {}

function! s:Excel.Book.new(file)
    let book = deepcopy(self)
    let book._file = a:file
    return book
endfunction

function! s:Excel.Book.get_filename()
    return self._file
endfunction

function! s:Excel.Book.get_lines()
    return s:com_get_lines(self._file)
endfunction


function! s:com_get_lines(file)
    let file = substitute(a:file, '/', '\', 'g')
    let csvfile = substitute(tempname(), '/', '\', 'g')
    let output = s:system(['cscript', s:WSH_SCRIPT, 'csvout', file, csvfile])
    if v:shell_error
        throw 'excel: Excel could not export csv file: ' . iconv(output, 'cp932', &enc)
    endif
    let csvlines = readfile(csvfile)
    let lines = []
    for csvline in csvlines
        let cols = s:parse_csvline(csvline)
        call add(lines, join(cols, "\t"))
    endfor
    return lines
endfunction

function! s:parse_csvline(line)
    " TODO: Add to Vital.Text.CSV
    let line = a:line
    let cols = []
    let rx_rest = '\%(,\|$\)\(.*\)'
    let rx_quotecol = '^"\(\|.\{-}\%(""\)*\)"' . rx_rest
    let rx_nonquotecol = '^\([^,]*\)' . rx_rest
    while line !=# ''
        if line[0] ==# '"'
            let m = matchlist(line, rx_quotecol)
        else
            let m = matchlist(line, rx_nonquotecol)
        endif
        if empty(m)
            throw 'failed to parse'
        endif
        call add(cols, m[1])
        let line = m[2]
    endwhile
    return cols
endfunction

function! s:error(msg)
    call s:echomsg('ErrorMsg', a:msg)
endfunction

function! s:echomsg(hl, msg)
    execute 'echohl' a:hl
    try
        echomsg a:msg
    finally
        echohl None
    endtry
endfunction

" Windows:
" Using :!start , execute program without via cmd.exe.
" Execute 'args' with 'noshellslash'
" keep special characters from unwanted expansion.
" (see :help shellescape())
"
" Unix:
" using :! , execute program in the background by shell.
function! s:system(args)
  if s:IS_MSWIN
    let shellslash = &l:shellslash
    setlocal noshellslash
  endif
  try
    let cmdline = join(map(a:args, 'shellescape(v:val)'), ' ')
    return system(cmdline)
  finally
    if s:IS_MSWIN
      let &l:shellslash = shellslash
    endif
  endtry
  return ''
endfunction


" Restore 'cpoptions' {{{
let &cpo = s:save_cpo
" }}}
