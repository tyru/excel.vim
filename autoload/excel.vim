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
    augroup excel
        autocmd!
        autocmd InsertEnter * call s:on_insert_enter()
    augroup END
    nnoremap <silent><buffer> <Plug>(excel-move-h) :<C-u>call <SID>map_move('h')<CR>
    nnoremap <silent><buffer> <Plug>(excel-move-j) :<C-u>call <SID>map_move('j')<CR>
    nnoremap <silent><buffer> <Plug>(excel-move-k) :<C-u>call <SID>map_move('k')<CR>
    nnoremap <silent><buffer> <Plug>(excel-move-l) :<C-u>call <SID>map_move('l')<CR>
    " nnoremap <silent><buffer> <Plug>(excel-edit-cell) :<C-u>call <SID>map_edit_cell()<CR>
    nmap <silent><buffer> h <Plug>(excel-move-h)
    nmap <silent><buffer> j <Plug>(excel-move-j)
    nmap <silent><buffer> k <Plug>(excel-move-k)
    nmap <silent><buffer> l <Plug>(excel-move-l)
    " nmap <silent><buffer> i <Plug>(excel-edit-cell)
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

    " Put cursor on A1 cell.
    " TODO: Get beginning cursor pos in excel book.
    call cursor(1, 1)
endfunction

function! s:on_insert_enter()
    " TODO: Change 'Cell.Text' to 'Cell.Value'.
endfunction

function! s:map_move(cmd)
    let [x, y] = s:get_cell_pos()
    let [x, y] = get({
    \   'h': [x - 1, y],
    \   'j': [x, y + 1],
    \   'k': [x, y - 1],
    \   'l': [x + 1, y],
    \}, a:cmd, [-1, -1])
    call s:set_cell_pos(x, y)
endfunction

function! s:get_cell_pos()
    let x = len(substitute(getline('.')[: col('.') - 1], '[^\t]', '', 'g'))
    let y = line('.') - 1
    return [x, y]
endfunction

function! s:set_cell_pos(x, y)
    if a:y < 0 || a:y >= line('$')
        return
    endif
    let lnum = a:y + 1
    let most_right_cell = len(substitute(getline(lnum), '[^\t]', '', 'g'))
    if a:x < 0 || a:x > most_right_cell
        return
    endif

    if a:x is 0
        let col = 1
    else
        let cols = split(getline(lnum), '\([^\t]*\zs\%(\t\|$\)\)', 1)
        " Remove empty string at the end of list.
        let cols = cols[: -2]
        if empty(cols)
            let col = 1
        else
            let col = len(join(cols[: a:x - 1], "\t")) + 2
        endif
    endif
    call cursor(lnum, col)
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
        call map(cols, 'substitute(v:val, "\\t", " ", "g")')
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
        if line[0] ==# '"'
            let m[1] = substitute(m[1], '""', '"', 'g')
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
