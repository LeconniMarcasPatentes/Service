from docx import Document

doc = Document('test.docx')
for paragraph in doc.paragraphs:
    t = paragraph.text
    t = t.split()
    
    # print( 't: ', t ) 
    # print( ' ' )
    
    if '[NOME_CONTRATANTE]' in t:        
        paragraph.text = paragraph.text.replace( '[NOME_CONTRATANTE]', 'Leconni Marcas & Patentes' ).upper( )
    else:
        print( 'Variavel não encontrada!' )

doc.save('test1.docx')
