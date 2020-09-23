.MODEL small
.STACK 100H
.DATA
       HelloMesg        db      10,13,'HOLA A TODOS!!!! QUE ONDA !?!?!?!',10,13,'$'
.CODE
START:
       mov       ax,@data
       mov       ds,ax
       mov       es,ax

       mov dx,OFFSET HelloMesg  ; offset of the text string
       mov ah,9                 ; print string function number
       int 21h                  ; dos call
       mov ah,4ch               ; terminate function number
       int 21h                  ; dos call

END

