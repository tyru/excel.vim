" vim:foldmethod=marker:fen:
scriptencoding utf-8

" Load Once {{{
if get(g:, 'loaded_excel', 0) || &cp
    finish
endif
let g:loaded_excel = 1
" }}}
" Saving 'cpoptions' {{{
let s:save_cpo = &cpo
set cpo&vim
" }}}



" Requirements {{{

let s:is_windows = has('win16') || has('win32') || has('win64') || has('win95')
if !s:is_windows
    echohl ErrorMsg
    echomsg 'error: your environment is not supported.'
    echohl None
    finish
endif
" if !executable('excel')
"     echohl ErrorMsg
"     echomsg 'error: Excel is not installed.'
"     echohl None
"     finish
" endif
" }}}



command! -nargs=1 -bang -complete=file
\   ExcelEdit
\   call excel#edit(<f-args>, <bang>0)



" Restore 'cpoptions' {{{
let &cpo = s:save_cpo
" }}}
