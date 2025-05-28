from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_UNDERLINE
from docx.shared import Pt, Inches, Cm
import requests
from io import BytesIO
from datetime import datetime
import qrcode
from io import BytesIO

def generar_y_insertar_qr(doc, link, tamaño_cm=4, alineacion='center', margen_superior=0.5):
    """
    Genera un código QR y lo inserta en un documento Word
    
    Args:
        doc (Document): Objeto Document de python-docx
        link (str): URL o texto para el código QR
        tamaño_cm (float): Tamaño del QR en centímetros
        alineacion (str): 'left', 'center' o 'right'
        margen_superior (float): Margen superior en cm
    
    Returns:
        None
    """
    try:
        # Generar QR en memoria
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(link)
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        
        # Guardar imagen en memoria
        img_bytes = BytesIO()
        img.save(img_bytes, format='PNG')
        img_bytes.seek(0)
        
        # Insertar en documento Word
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Cm(margen_superior)
        
        if alineacion == 'left':
            paragraph.alignment = 0  # WD_ALIGN_PARAGRAPH.LEFT
        elif alineacion == 'right':
            paragraph.alignment = 2  # WD_ALIGN_PARAGRAPH.RIGHT
        else:
            paragraph.alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER
        
        run = paragraph.add_run()
        run.add_picture(img_bytes, width=Cm(tamaño_cm))
        
        return True
    
    except Exception as e:
        print(f"Error al insertar QR: {str(e)}")
        return False

def limpiar_documento(path):
    doc = Document(path)

    # Eliminar todos los párrafos
    for _ in range(len(doc.paragraphs)):
        p = doc.paragraphs[0]
        p._element.getparent().remove(p._element)

    # Eliminar todas las tablas (si las hay)
    for _ in range(len(doc.tables)):
        t = doc.tables[0]
        t._element.getparent().remove(t._element)

    return doc

def  FORM_DATOS_NUEVOS_PARA_TRABAJADOR(datos_cliente):
    def agregar_item(doc, etiqueta, valor,subrayado=0):
        p = doc.add_paragraph(style='List Bullet')  # Estilo viñeta
        p.add_run(f"{etiqueta}: ").bold = True

        p.add_run(valor)
        run = p.runs[0]
        run.font.size = Pt(12)

    def agregar_si_no(doc, campo, valor=None):
            p = doc.add_paragraph(style='List Bullet')
            
            # Nombre del campo en negrita
            p.add_run(f"{campo}:").bold = True
            
            # Tabulación para alinear (ajustar según necesidad)
            p.add_run("\t" * 3)  # 3 tabulaciones
            
            # Opción SI (siempre en negrita)
            si = p.add_run("SI")
            si.bold = True
            if valor.lower()=='si':
                si.underline = WD_UNDERLINE.SINGLE  # Subrayar solo si es True
            
            # Espacio entre opciones (2 tabulaciones)
            p.add_run("\t" * 2)
            
            # Opción NO (siempre en negrita)
            no = p.add_run("NO")
            no.bold = True
            if valor.lower()=='no':
                no.underline = WD_UNDERLINE.SINGLE  # Subrayar solo si es False

    def datos_ubicacion(doc, etiqueta, valor):
        p = doc.add_paragraph(style='List Bullet')  # Estilo viñeta
        p.add_run(f"{etiqueta}: ")

        p.add_run(valor)
        run = p.runs[0]
        run.font.size = Pt(11)

    def get_google_drive_image(drive_url):
        """Convierte un enlace compartido de Google Drive a enlace directo"""
        if '/open?id=' in drive_url:
            # Para enlaces del formato: https://drive.google.com/open?id=FILE_ID
            file_id = drive_url.split('/open?id=')[1].split('&')[0]
        else:
            # Para enlaces del formato: https://drive.google.com/file/d/FILE_ID/
            file_id = drive_url.split('/d/')[1].split('/')[0]
        
        return f"https://drive.google.com/uc?export=view&id={file_id}"

    def add_image_from_drive(doc, drive_url,ubicacion_url, width=Inches(4)):
        try:
            # Obtener enlace directo
            direct_url = get_google_drive_image(drive_url)
            
            # Descargar imagen
            response = requests.get(direct_url)
            img_data = BytesIO(response.content)
            
            # Agregar al documento
            doc.add_picture(img_data, width=width)
            doc.add_paragraph(f": {ubicacion_url}", style='Caption')
            
        except Exception as e:
            print(f"Error al insertar imagen: {e}")
            doc.add_paragraph(f"[Imagen no disponible: {drive_url}]")
    document = Document()
    estilo_normal = document.styles["Normal"]
    estilo_normal.font.name = "Arial"
    # Obtener la fecha actual
    fecha_actual = datetime.now().strftime('%d/%m/%Y')

    fecha = document.add_paragraph()
    fecha.add_run('Fecha: ').bold=True
    fecha.add_run(fecha_actual).italic = True
    run = fecha.runs[0]
    run.font.size = Pt(11)
    run.italic = True
    fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    #formularios

    title = document.add_paragraph()
    title.add_run('DATOS DEL TRABAJADOR').bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(14)
    run.italic = True

    document.add_paragraph() #espacio

    def agregar_item(doc, etiqueta, valor,subrayado=0):
        p = doc.add_paragraph(style='List Bullet')  # Estilo viñeta
        p.add_run(f"{etiqueta}: ").bold = True

        p.add_run(valor)
        run = p.runs[0]
        run.font.size = Pt(12)
            

    agregar_item(document,'NOMBRE Y APELLIDO ',datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'])
    agregar_item(document,'ESTADO CIVIL',datos_cliente['Estado Civil como esta en tu cedula'])
    agregar_item(document,'NACIONALIDAD',datos_cliente['Nacionalidad'])
    agregar_item(document,'NÚMERO DE C.I.',datos_cliente['Numero de Cedula'])
    agregar_item(document,'CIUDAD',datos_cliente['Ciudad'])
    agregar_item(document,'BARRIO',datos_cliente['Barrio'])
    agregar_item(document,'DIRECCION - CALLES',datos_cliente['Direccion Particular, Calles, Numero de casa'])
    agregar_item(document,'TELEFONO',datos_cliente['Telefono de contacto personal'])
    agregar_item(document,'EMPRESA EN QUE TRABAJA O TRABAJO',datos_cliente['Empresa en la que trabajo <Razon Social>'])
    agregar_item(document,'DIRECCION DE LA EMPRESA',datos_cliente['Direccion de la Empresa'])
    agregar_item(document,'RUC DE LA EMPRESA',datos_cliente['Ruc de la empresa'])
    agregar_item(document,'FECHA DE INGRESO',datos_cliente['Fecha de ingreso'])
    agregar_item(document,'FECHA DE DESPIDO',datos_cliente['Fecha de Despido'])
    agregar_item(document,'QUE DIAS TRABAJA',datos_cliente['JORNADA LABORAL. Como es o era tu Jornada Laboral? Lunes a Viernes, Lunes a Lunes?'])
    agregar_item(document,'HORARIO DE TRABAJO',datos_cliente['HORARIO DE TRABAJO. Como era tu Horario que Cumplias? Ej 8.00 a 18.00'])
    agregar_item(document,'SALARIO',datos_cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?'])
    agregar_item(document,'IPS',datos_cliente['IPS'])
    agregar_item(document,'BONIFICAION FAMILIAR (CUANTOS HIJOS)',datos_cliente['Bonficacion familiar por Cuantos hijos si la respuesta fue SI'])
    agregar_item(document,'TAREA REALIZADA (DESCRIPCION)',datos_cliente['Describe las tareas o funciones que desempenabas en el lugar de trabajo.'])
    agregar_item(document,'MOTIVO DEL DESPIDO','Chau')
    agregar_item(document,'QUIEN LO DESPIDIO Y POR QUE MEDIO',datos_cliente['MOTIVO DE DESPIDO. Cuentanos como se dio la situacion.'])

    agregar_item(document,'LE DEBEN SALARIOS',datos_cliente['SALARIOS PENDIENTES. Te deben Salarios, Cuanto de cuantos dias o meses?'])
    agregar_item(document,'PAGO DE SALARIOS (EFECTIVO U OTRO)',datos_cliente['PAGOS DE SALARIOS. Como recibias los pagos de salario o jornales?. Efectivo, Trasferencia, giros? via que banco?'])
    agregar_item(document,'LE ENTREGABAN RECIBOS DE SALARIOS','')
    agregar_item(document,'LE DEBEN VACACIONES',datos_cliente['VACACIONES. Salias o tenias vacaciones? Te deben vacaciones?'])
    agregar_item(document,'LE DEBEN AGUINALDO',datos_cliente['AGUINALDO. Recibias Aguinaldo?. Te pagaban o te debe?'])

    agregar_si_no(document,'CONTABA CON CONTRATO','no')

    agregar_si_no(document,'LE OFRECIERON SU LIQUIDACION',datos_cliente['CONTRATO DE TRABAJO. Tenias contrato de Trabajo Firmado'])
    agregar_si_no(document,'FIRMO DOCUMENTOS EN BLANCO',datos_cliente['Firmaste en algun momento algun Documento en blanco o pagare?'])
    agregar_item(document,'OBSERVACIONES',datos_cliente['Alguna informacion adicional que deseas agregar?'])

    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE) #SGTE HOJA


    #ubicaion del actor
    titulo_actor = document.add_paragraph()
    titulo_actor.add_run('CROQUIS DE UBICACION DEL DOMICILIO DEL ACTOR').bold = True
    titulo_actor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo_actor.runs[0]
    run.font.size = Pt(11)
    run.italic = True



    datos_ubicacion(document,'NOMBRE Y APELLIDO',datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'])
    datos_ubicacion(document,'DIRECCION',datos_cliente['Direccion Particular, Calles, Numero de casa'])
    datos_ubicacion(document,'CIUDAD',datos_cliente['Ciudad'])
    datos_ubicacion(document,'TELEFONO',datos_cliente['Telefono de contacto personal'])

    document.add_paragraph()  #espacio

    add_image_from_drive(document,datos_cliente['Adjunta Imagen de la Ubicacion de Google Maps de casa Trabajador.'],datos_cliente['Ubicacion de tu casa. Copia el link de la ubicacion de google'])

    generar_y_insertar_qr(document,datos_cliente['Ubicacion de tu casa. Copia el link de la ubicacion de google'])
    # Crear párrafo para la firma alineado a la derecha
    firma = document.add_paragraph()
    firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Agregar línea horizontal (subrayado)
    run = firma.add_run("_________________________")
    run.font.size = Pt(12)

    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)#SGTE HOJA

    #ubicacion del demandado
    titulo_demandado = document.add_paragraph()
    titulo_demandado.add_run('CROQUIS DE UBICACION DEL DOMICILIO DEL DEMANDADO').bold = True
    titulo_demandado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo_demandado.runs[0]
    run.font.size = Pt(11)
    run.italic = True

    add_image_from_drive(document,datos_cliente['Adjunta Imagen de la Ubicacion de Google Maps de la empresa'],datos_cliente['Ubicacion de la empresa'])
    generar_y_insertar_qr(document,datos_cliente['Ubicacion de la empresa'])

    document.add_paragraph() #salto de linea
    document.add_paragraph() #salto de linea
    document.add_paragraph() #salto de linea

    agregar_item(document,"Entrevista realizada por",datos_cliente['Entrevista realizada por'])

    # Agregar texto "FIRMA" debajo
    firma = document.add_paragraph()
    firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = firma.add_run("FIRMA")
    run.font.size = Pt(11)
    
    document.save(f'Generado/Formulario_Datos_{datos_cliente['Numero de Cedula']}.docx')


def obtener_fecha_formateada():
    # Diccionario de meses en español
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    
    # Obtener fecha actual
    hoy = datetime.now()
    dia = hoy.day
    mes_num = hoy.month
    año = hoy.year
    
    # Formatear la fecha
    fecha_formateada = f"{dia} días del mes {meses[mes_num]} del año {año}"
    
    return fecha_formateada

from docx import Document

def agregar_parrafo_con_negrita(doc, text, bold_phrases):
    """
    Agrega un párrafo al documento con ciertas palabras/frases en negrita.
    
    :param doc: Objeto Document de python-docx.
    :param text: Texto completo del párrafo.
    :param bold_phrases: Lista de palabras o frases que deben ir en negrita.
    """
    paragraph = doc.add_paragraph()
    current_pos = 0

    # Ordenar por longitud inversa para evitar problemas de solapamiento
    sorted_bolds = sorted(bold_phrases, key=len, reverse=True)
    
    while current_pos < len(text):
        match = None
        match_start = len(text)
        
        for phrase in sorted_bolds:
            idx = text.find(phrase, current_pos)
            if idx != -1 and idx < match_start:
                match = phrase
                match_start = idx
        
        if match:
            if match_start > current_pos:
                # Texto normal antes de la frase en negrita
                paragraph.add_run(text[current_pos:match_start])
            # Frase en negrita
            bold_run = paragraph.add_run(match)
            bold_run.bold = True
            current_pos = match_start + len(match)
        else:
            # Resto del texto sin negrita
            paragraph.add_run(text[current_pos:])
            break


def Carta_Poder(datos_clientes):
    document = limpiar_documento('Plantilla/Carta_de_Poder.docx')
    style = document.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.sizwe = Pt(11)

    titulo = document.add_paragraph("CARTA PODER")
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #cuerpo
    cuerpo = document.add_paragraph()
    cuerpo.add_run('En la ciudad de Asunción, a los ')
    cuerpo.add_run(obtener_fecha_formateada()+', comparecen ')


    if(datos_clientes['Sexo']=='Femenino'):
        cuerpo.add_run(f'la Sra. {datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.'].upper()} ').bold=True
    else:
        cuerpo.add_run(f'El Sr. {datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.'].upper()} ').bold=True

    cuerpo.add_run('paraguayo/a, mayor de edad, con ')
    cuerpo.add_run(f'C.I. N° {datos_clientes['Numero de Cedula']}, ').bold=True
    cuerpo.add_run('soltero/a, empleado con domicilio en las calles ')
    cuerpo.add_run(f'{datos_clientes['Direccion Particular, Calles, Numero de casa']}, Barrio {datos_clientes['Barrio']} de la ciudad de {datos_clientes['Ciudad']}; ')
    cuerpo.add_run('Y DICE: ').bold=True
    cuerpo.add_run('Que confieren  mandato a los  abogados, ')
    cuerpo.add_run('JUAN JOSE BERNIS, ').bold=True
    cuerpo.add_run('con matrícula ')
    cuerpo.add_run('Nº 18.500, ESTELA NOGUERA, ').bold=True
    cuerpo.add_run('con matrícula ')
    cuerpo.add_run('N° 50.511 Y MARCELO LEITE, ').bold=True
    cuerpo.add_run('con matrícula ')
    cuerpo.add_run('N° 53.693, ').bold=True
    cuerpo.add_run('para que me representen, en el carácter invocado, y luego de aceptarlo, y quien acepte, ante los Jueces de Primera Instancia  del Trabajo, Jueces de Primera Instancia en lo Civil, Comercial , Laboral y Tutelar del Menor, Cámara de Apelación en lo Laboral, Cámara de Apelación en lo Civil, Comercial, Laboral, y Tutelar del Menor, Jueces Electorales, Tribunal Electoral, ante la Corte Suprema de Justicia, la Dirección General del Trabajo,  y ante cualquier autoridad del mismo fuero, interpongan y  contesten amparo, recurran medidas cautelares, soliciten tales medidas, sea como actor o demandado, que realicen todos los actos concernientes  a las gestiones autorizadas por las Leyes del Trabajo y/o las nuevas leyes laborales, administrativas promulgadas o a ser promulgadas , por las disposiciones que reglan la Dirección General del  Trabajo y  el INSTITUTO DE PREVISION SOCIAL, por las demás reglamentaciones que rigen el país, y que sean aplicables a las cuestiones y conflictos laborales, facultándolos para ejercitar este mandato  para demandar, contestar demanda y/o reconvenir contra ')
    cuerpo.add_run(f'{datos_clientes['Empresa en la que trabajo <Razon Social>'].upper()} con RUC Nº {datos_clientes['Ruc de la empresa']}, ').bold=True
    cuerpo.add_run(f'con ubicación en las calles {datos_clientes['Direccion de la Empresa']} de la ciudad de {datos_clientes['Ciudad de la empresa']} ')
    otorga = cuerpo.add_run('OTORGA')
    otorga.bold=True
    otorga.underline=True
    cuerpo.add_run(' mandato a los abogados de la matrícula para que lo representen en todas las gestiones administrativas y judiciales, interponiendo la demanda, contestando la demanda, o reconviniendo, recurriendo por inconstitucionalidad ante la Corte Suprema de Justicia, ofreciendo pruebas, testimoniales, absuelvan posiciones en nombre de los mismos, confesarías, pidan medidas precautelares para impedir simulaciones o fraudes,  exijan se trabe embargos sobre bienes del demandado, soliciten medidas preliminares, apelan, interponga todo tipo de recursos, adelanten pruebas, substituyan mandato  en forma parcial o total a favor de otro, nombren al perito de parte,  acepten pruebas o impugnaciones  en su caso,  y de la demandada, opongan excepciones, prescripciones, , denuncien la falta de acción, de personería, etc., recusen a los Jueces, pidan inhibiciones de funcionarios  comprendidos dentro de las generales de la Ley, delimiten pruebas a ofrecer, presenten al mandante en las audiencias fijadas por el Juez, desistir de la acción y del derecho, denuncien en jurisdicción internacional  la falta de cumplimiento de tratados o convenios suscritos o ratificados por la República del Paraguay, se presente en queja ante los organismos internacionales por lesiones a la libertad sindical, negociación colectiva, pidan informaciones a la administración central, por medio del Habeas data u otra forma de petitorio , desistan de testigos y otras formas de prueba, declinen de instancias, hagan denuncias penales, instrumenten por reenvío la aplicación del Código Civil,  Código de Procedimientos Civil, en todo cuanto beneficie al trabajador, soliciten el libramiento de exhortos u oficios a otros Jueces de extraña jurisdicción, nombren árbitros y/o arbitradores, sometan el caso a los mismos, acepten  y/o rechacen el veredicto de los nombrados, acusen de nulidad los actos procesales, y realicen cuanto más actos procesales sean pertinentes y toda presentación mía, bajo patrocinio de abogado, revocará mandato si expresamente no mencionara ‘’sin revocar mandato’’, entendiéndose que este mandato es oneroso.- El mandato otorgado en este instrumento quedará perfeccionado, formalizado, aceptado, e instrumentado recién al momento en el que el abogado designado se presente ante un juzgado u órgano extrajudicial y solicite intervención y reconocimiento de personería, y solamente comprometerá a quien aceptó el mandato, excluyendo de toda responsabilidad a quienes no aceptaron el mandato, aunque hayan sido mencionados en esta carta poder. En  prueba de conformidad, suscriben los mismos  el presente mandato y/o carta poder, que deberá ser aceptado individualmente.')
    
    
    document.save(f'Generado/Carta_de_Poder_{datos_clientes['Numero de Cedula']}.docx')

def Carta_Compromiso(datos_clientes):
    doc = limpiar_documento('Plantilla/Carta_Compromiso.docx') #crea un documento en blanco con la plantill Carta_Compromiso.docx
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Georgia'
    font.size = Pt(12)

    #titulo
 
    
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER    
    run_titulo = titulo.add_run("NOTA DE COMPROMISO")
    run_titulo.underline = True
    run_titulo.bold = True
    
    #cuerpo
    p1 = doc.add_paragraph()
    p1.add_run(f'En la ciudad de Asuncion a los {obtener_fecha_formateada()}.')
    
    p2=f'''Yo, {datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.'].upper()} con C.I Nº {datos_clientes['Numero de Cedula']} con domicilio en las calles {datos_clientes['Direccion Particular, Calles, Numero de casa']} de la Ciudad de {datos_clientes['Ciudad']},por medio de la presente, me comprometo formalmente a asistir a las instalaciones del Estudio Jurídico BERNIS ALLEGRETTI ubicado en las calles Milano 282 C/ Chile de la Ciudad de Asunción, con el fin de hacer el seguimiento correspondiente al juicio a ser iniciado en contra de la empresa {datos_clientes['Empresa en la que trabajo <Razon Social>'].upper()}, con RUC Nº {datos_clientes['Ruc de la empresa']}'''
    agregar_parrafo_con_negrita(doc,p2,[f'{datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.'].upper()} con C.I Nº {datos_clientes['Numero de Cedula']}',
                                    'BERNIS ALLEGRETTI',
                                    f'{datos_clientes['Empresa en la que trabajo <Razon Social>'].upper()}, con RUC Nº {datos_clientes['Ruc de la empresa']}'])
    
    p3='\tQue, conforme a lo anterior, reconozco que es mi responsabilidad mantener el seguimiento adecuado de mi caso, asistir a las citas y proporcionar la información o documentación requerida por los abogados.'
    agregar_parrafo_con_negrita(doc,p3,['\tQue,'])
    
    p4 = '\tQue, el compromiso de asistencia será de manera periódica y según lo establecido por el abogado JUAN JOSE BERNIS ALLEGRETTI encargado de mi caso, a fin de mantener una comunicación continua y eficaz sobre la demanda.'
    agregar_parrafo_con_negrita(doc,p4,['\tQue,'])
    
    p5 = '\tQue, reconozco que este seguimiento es fundamental para el adecuado desarrollo de mi caso, y me comprometo a cumplir con las citas y solicitudes que el equipo jurídico considere necesarias con el fin de evitar la PERENCION DE LA INSTANCIA en la causa, el cual se debe a la inactividad procesal por el lapso de 3 (tres) meses.'
    agregar_parrafo_con_negrita(doc,p5,['\tQue,'])
    
    p6='\tQue, de igual manera, dejo claro que el abogado JUAN JOSE BERNIS ALLEGRETTI con Mat. Nº 18.500 y el Estudio Jurídico BERNIS ALLEGRETTI no asumen responsabilidad alguna en caso de que no cumpla con mis compromisos de asistencia o si no tomo las decisiones necesarias relacionadas con el proceso judicial que me compite. Eximo de responsabilidad al abogado y al Estudio Jurídico por cualquier consecuencia derivada de mi accionar.'
    agregar_parrafo_con_negrita(doc,p6,['\tQue,',
                                        'JUAN JOSE BERNIS ALLEGRETTI',
                                        'Mat. Nº 18.500',
                                        'BERNIS ALLEGRETTI'])
    
    p7='\tQue, a fin de corroborar mi asistencia firmare el acta de asistencia del Estudio Jurídico.-'
    agregar_parrafo_con_negrita(doc,p7,['\tQue'])
    
    p8 = '\tQue, además reconozco no poseer documentos respaldatorios que demuestren mi relación por lo cual he sido advertido de las posibilidades en laboral en el juicio.-'
    agregar_parrafo_con_negrita(doc,p8,[])
    
    p9 = '\tSin más en particular, firmo al pie de la presente nota en señal de conformidad. –'
    agregar_parrafo_con_negrita(doc,p9,[])
    
    p10 = 'FIRMA:\nACLARACION:\nC.I. N°:'
    agregar_parrafo_con_negrita(doc,p10,['FIRMA:\nACLARACION:\nC.I. N°:'])
    
    doc.save(f'Generado/Carta_Compromiso_{datos_clientes['Numero de Cedula']}.docx')
    
    
def Desistimiento_de_reuncia(datos_clientes):
    doc = limpiar_documento('Plantilla/Desistimiento_de_Renuncia.docx')
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    #fecha
    fecha = doc.add_paragraph()
    fecha.add_run('Asunción,___de________del año 202__')
    fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph() #salto de linea
    
    p1 = doc.add_paragraph()
    p1.add_run('SUPERINTENDENCIA DE LA CORTE SUPREMA DE JUSTICIA').bold=True
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p2 = doc.add_paragraph()
    p2.add_run('Alonso e/ Testanova de la Ciudad de Asunción.')
    agregar_parrafo_con_negrita(doc,'Asunto: Desistimiento de Denuncia. -',['Asunto:']) #p3
    p4=doc.add_paragraph()
    p4.add_run('Estimados Señores:').bold=True

    doc.add_paragraph() #salto de linea
    
    p5 = f'\tYo, {datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.']} con C.I Nº {datos_clientes['Numero de Cedula']}, por medio de la presente me dirijo a ustedes en mi calidad de denunciante en el caso relacionado con el abogado Juan José Bernis Allegretti con Mat. Nº 18.500. –'
    agregar_parrafo_con_negrita(doc,p5,[f'{datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.']} con C.I Nº {datos_clientes['Numero de Cedula']},'])
    
    p6 = '\tMediante esta nota, deseo formalizar mi desistimiento de la denuncia presentada en contra del profesional Abg. Juan Bernis,'
    agregar_parrafo_con_negrita(doc,p6,[])
    
    p7 = '\tEn consecuencia, reconozco que la denuncia carece de fundamento y no procede su prosecución. –'
    agregar_parrafo_con_negrita(doc,p7,[])
    
    p8 = '\tAgradezco la atención brindada a este asunto y lamento cualquier inconveniente que mi denuncia haya podido ocasionar. –'
    agregar_parrafo_con_negrita(doc,p8,[])
    
    p9= '\tSin otro particular, me despido atentamente. –'
    agregar_parrafo_con_negrita(doc,p9,[])
    
    doc.save(f'Generado/Desistimiento_de_Renuncia_{datos_clientes['Numero de Cedula']}.docx')



def Nota_de_Renuncia(datos_clientes):
    doc = limpiar_documento('Plantilla/Nota_de_Renuncia.docx')
    #estilos
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Georgia'
    font.size = Pt(12)
    
    doc.add_paragraph() #salto de linea
    doc.add_paragraph() #salto de linea
    #fecha
    fecha = doc.add_paragraph()
    fecha.add_run('Asunción,___de________del año 202__')
    fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph() #salto de linea
    
    p1 = doc.add_paragraph()
    run = p1.add_run(f'SR/A. {datos_clientes['Nombres y Apellidos completos como esta en tu Cedula.'].upper()} con C.I Nº {datos_clientes['Numero de Cedula']}')
    run.font.size = Pt(11)
    run.bold = True
    
    agregar_parrafo_con_negrita(doc,'Objeto: Renunciar al Mandato. –',['objeto:']) #p2
    
    p3 = f'\tQue, por medio de la presente, me dirijo a usted para comunicarle mi decisión de renunciar al mandato profesional que me confirió con la finalidad de promover demanda laboral en contra de {datos_clientes['Empresa en la que trabajo <Razon Social>']}, con RUC Nº {datos_clientes['Ruc de la empresa']} motivos particulares.'
    agregar_parrafo_con_negrita(doc,p3,[f'{datos_clientes['Empresa en la que trabajo <Razon Social>'].upper()}, con RUC Nº {datos_clientes['Ruc de la empresa']}'])
    
    p4 = '\tQue, le informo que, a partir de la fecha, ceso en mi representación legal en el asunto mencionado.'
    agregar_parrafo_con_negrita(doc,p4,['\tQue,'])
    
    p5 = '\tQue, asimismo, intimo nombre un nuevo representante en un plazo no mayor de 5 (CINCO) DIAS, para que se haga cargo del caso y así evitar perjuicios, no teniendo ninguna responsabilidad legal mi persona como profesional desde esta notificación. –'
    agregar_parrafo_con_negrita(doc,p5,['\tQue,',
                                        '5 (CINCO) DIAS,'])
    
    p6 = '\tQue, quedo a su disposición para coordinar la entrega de documentos y cualquier otra gestión necesaria para facilitar la transición. –'
    agregar_parrafo_con_negrita(doc,p6,['\tQue,'])
    
    p7 = f'\tQue, agradezco la confianza depositada en mí durante este tiempo y lamento cualquier inconveniente que esta decisión pueda causarle. –'
    agregar_parrafo_con_negrita(doc,p7,['\tQue,'])
    
    doc.add_paragraph() #salto de linea
    
    doc.add_paragraph('Atentamente, ')

    doc.add_paragraph() #salto de linea
    doc.add_paragraph() #salto de linea

    abogado = doc.add_paragraph()
    abogado.add_run('JUAN JOSE BERNIS ALLEGRETTI').bold=True
    abogado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    matricula = doc.add_paragraph()
    matricula.add_run('MAT. Nº 18.500').bold=True
    matricula.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.save(f'Generado/Nota_de_Renuncia_{datos_clientes['Numero de Cedula']}.docx')
    
def documento_demanda(datos_cliente):
    doc = limpiar_documento('Plantilla/documento_demanda.docx')
    #estilos
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    
    #tabla
    tabla = doc.add_table(rows=4, cols=2)


    # Datos para la tabla 
    datos = [
        ["OBJETO:", "DEMANDA POR PAGO DE APORTES AL INSTITUTO DE PREVISIÓN SOCIAL E INDEMNZIACIÓN POR DAÑO MORAL ANTE EL NO PAGO POR EL SALARIO REAL PERCIBIDO, Y OTROS BENEFICIOS LABORALES. -"],
        ["ACTOR:", datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper()],
        ["DEMANDADA:", datos_cliente['Empresa en la que trabajo <Razon Social>'].upper()],
        ["SEÑOR/A JUEZ:", ""]
    ]

    # Llenar la tabla con los datos
    for i, (etiqueta, valor) in enumerate(datos):
        # Celda izquierda (etiqueta)
        celda_izq = tabla.cell(i, 0)
        celda_izq.text = etiqueta
        celda_izq.paragraphs[0].runs[0].font.bold = True
        
        # Celda derecha (valor)
        celda_der = tabla.cell(i, 1)
        celda_der.text = valor
        
    doc.add_paragraph() #salto de linea
    
    p1 = f'\t\tJUAN JOSÉ BERNIS, Abogado de la Matrícula Nº 18.500, en nombre y representación de 1) {datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper()}, de nacionalidad {datos_cliente['Nacionalidad'].lower()}/a, {datos_cliente['Estado Civil como esta en tu cedula']}/a, mayor de edad, empleada, con C.I. Nº {datos_cliente['Numero de Cedula']}, con domicilio real en calle {datos_cliente['Direccion Particular, Calles, Numero de casa']}, Barrio {datos_cliente['Barrio']}, de la Ciudad de {datos_cliente['Ciudad'].split()[0]}; constituyendo domicilio procesal en calle Milano N° 282 esq. Chile, de la Ciudad de Asunción, a V.S., respetuosamente, digo:'
    agregar_parrafo_con_negrita(doc,p1,['JUAN JOSÉ BERNIS,',f'1) {datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper()},'])
    
    p2 = f'\t\tPERSONERIA: Adjunto carta poder de la Señora 1) 1) {datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper()}, de nacionalidad {datos_cliente['Nacionalidad'].lower()}/a, {datos_cliente['Estado Civil como esta en tu cedula']}/a, mayor de edad, empleada, con C.I. Nº {datos_cliente['Numero de Cedula']}, con domicilio real en calle {datos_cliente['Direccion Particular, Calles, Numero de casa']}, Barrio {datos_cliente['Barrio']}, de la Ciudad de {datos_cliente['Ciudad'].split()[0]}; quien me otorga mandato para que intervenga en su nombre y representación, en todo asunto judicial o administrativo en que éste sea parte, en virtud de conflictos que se susciten con motivo de derechos emergentes de las leyes laborales.'
    agregar_parrafo_con_negrita(doc,p2,['PERSONERIA',f'1) {datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper()},'])
    
    p3 = f'\t\tCon suficiente poder y en cumplimiento del mismo, vengo a incoar demanda por INDEMNIZACIÓN POR DAÑO MORAL ANTE EL NO PAGO DEL SEGURO SOCIAL OBLIGATORIO POR EL SALARIO REAL PERCIBIDO, PAGO DE APORTES AL INSTITUTO DE PREVISION SOCIAL POR TODOS LOS AÑOS NO APORTADOS Y DIFERENCIA DE APORTES SOBRE EL SALARIO REAL PERCIBIDO, INDEMNIZACIÓN COMPENSATORIA Y COMPLEMENTARIA, HORAS EXTRAS, 30% POR NOCTURNIDAD, AGUINALDO Y VACACIONES CAUSADAS 2024 Y PROPORCIONALES 2025, PRE AVISO E INDEMNIZACIÓN, INTERESES DEL 3% MENSUAL DESDE EL NO PAGO DEL APORTE A LA SEGURIDAD SOCIAL HASTA EL EFECTIVO PAGO, MÁS COSTOS Y COSTAS contra {datos_cliente['Empresa en la que trabajo <Razon Social>'].upper()} RUC N° {datos_cliente['Ruc de la empresa']}, con domicilio en calle {datos_cliente['Direccion de la Empresa']}, de la Ciudad de {datos_cliente['Ciudad de la empresa']}. -'
    agregar_parrafo_con_negrita(doc,p3,['INDEMNIZACIÓN POR DAÑO MORAL ANTE EL NO PAGO DEL SEGURO SOCIAL OBLIGATORIO POR EL SALARIO REAL PERCIBIDO, PAGO DE APORTES AL INSTITUTO DE PREVISION SOCIAL POR TODOS LOS AÑOS NO APORTADOS Y DIFERENCIA DE APORTES SOBRE EL SALARIO REAL PERCIBIDO, INDEMNIZACIÓN COMPENSATORIA Y 	COMPLEMENTARIA, HORAS EXTRAS, 30% POR NOCTURNIDAD, AGUINALDO Y VACACIONES CAUSADAS 2024 Y PROPORCIONALES 2025, PRE AVISO E INDEMNIZACIÓN, INTERESES DEL 3% MENSUAL DESDE EL NO PAGO DEL APORTE A LA SEGURIDAD SOCIAL HASTA EL EFECTIVO PAGO, MÁS COSTOS Y COSTAS'])    
    
    doc.add_paragraph() #salto de linea
    doc.add_paragraph() #salto de linea
    
    titulo = doc.add_paragraph()
    titulo.add_run('HECHOS').bold=True
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitulo1 = doc.add_paragraph()
    subtitulo1.add_run('1.- FECHA DE INGRESO:').bold=True
    
    
    meses = {   
        '01': "enero", '02': "febrero", '03': "marzo", '04': "abril",
        '05': "mayo", '06': "junio", '07': "julio", '08': "agosto",
        '09': "septiembre", '10': "octubre", '11': "noviembre", '12': "diciembre"
    }
    
    fecha=datos_cliente['Fecha de ingreso'].split('/')
    p4 = f'El actor, ingresó a trabajar en típica relación de dependencia, trabajo continuo y permanente, en fecha {fecha[0]} de {meses[fecha[1]]} de {fecha[2]}, bajo la dependencia de la demandada.'
    agregar_parrafo_con_negrita(doc,p4,['ingresó',f'{fecha[0]} de {meses[fecha[1]]} de {fecha[2]},'])
    
    subtiulo2 = doc.add_paragraph()
    subtiulo2.add_run('2- TAREA REALIZADA:').bold=True
    
    p5 = f'El trabajador fue contratado para cumplir funciones como {datos_cliente['Describe las tareas o funciones que desempenabas en el lugar de trabajo.']}, bajo dependencia de las demandada.'
    agregar_parrafo_con_negrita(doc,p5,[])
    
    subtiulo3 = doc.add_paragraph()
    subtiulo3.add_run('3.- HORARIO DE TRABAJO:').bold=True
    
    p6 = f'LA actora cumplía un horario de trabajo de {datos_cliente['JORNADA LABORAL. Como es o era tu Jornada Laboral? Lunes a Viernes, Lunes a Lunes?']} de {datos_cliente['HORARIO DE TRABAJO. Como era tu Horario que Cumplias? Ej 8.00 a 18.00']}, SE RECLAMAN LAS HORAS EXTRAS DIURNAS Y NOCTURNAS POR DOCE MESES ANTES DEL DESPIDO.'
    agregar_parrafo_con_negrita(doc,p6,[])
    
    subtiulo4 = doc.add_paragraph()
    subtiulo4.add_run('4- SALARIO PERCIBIDO:').bold=True
    
    p7 = f'El trabajador percibía un salario promedio mensual de GS. {datos_cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?']}, NO SE LE ABONABA EL SEGURO SOCIAL OBLIGATORIO POR EL SALARIO REAL PERCIBIDO '
    agregar_parrafo_con_negrita(doc,p7,[])
    
    subtiulo5 = doc.add_paragraph()
    subtiulo5.add_run('5.- CONFIGURACIÓN DEL DESPIDO:').bold=True
    
    fecha_despido = datos_cliente['Fecha de Despido'].split('/')
    p8 = f'El TRABAJADOR FUE DESPEDIDO EN FECHA {fecha_despido[0]} de {meses[fecha_despido[1]]} DE {fecha_despido[2]}, sin que se le abone las indemnizaciones legales por despido injustificado. -'
    agregar_parrafo_con_negrita(doc,p8,[f'El TRABAJADOR FUE DESPEDIDO EN FECHA {fecha_despido[0]} de {meses[fecha_despido[1]]} DE {fecha_despido[2]}, sin que se le abone las indemnizaciones legales por despido injustificado. -'])
    
    p9 = f'6.- PAGO DE APORTES AL INSTITUTO DE PRECISIÓN SOCIAL DE LOS AÑOS {fecha[2]} hasta el {fecha_despido[2]} Y PAGO DE APORTES AL IPS SOBRE EL SALARIO REAL PERCIBIDO.'
    agregar_parrafo_con_negrita(doc,p9,['PAGO DE APORTES AL INSTITUTO DE PRECISIÓN SOCIAL'])
    
    p10 = f'Que la demandada no abono la seguridad social desde el inicio de la relación laboral POR EL SALARIO REAL PERCIBIDO DE GS {datos_cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?']}, habiéndole descontado siempre el 9% al trabajador para el supuesto aporte, DESDE YA, SOLICITO SE LIBRE OFICIO AL INSTITUTO DE PREVISIÓN SOCIAL A FIN DE QUE REMITA LA PLANILLA DE APORTES REALIZADO POR LA EMPRESA A FAVOR DEL TRABAJADOR. '
    agregar_parrafo_con_negrita(doc,p10,[f'POR EL SALARIO REAL PERCIBIDO DE GS {datos_cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?']}, habiéndole descontado siempre el 9% al trabajador para el supuesto aporte,'])
    
    p11 = 'Que, el actor, toma conocimiento de estos hechos en fecha ________________________, cuando ingresa a la página institucional del IPS, a consultar sus aportes, pero en el mismo consta que no tiene aportes a la seguridad social -'
    
    subtiulo6 = doc.add_paragraph()
    subtiulo6.add_run('DAÑO MORAL').bold=True
    
    p12 = f'SE RECLAMA UNA INDEMNIZACIÓN POR DAÑO MORAL por el no pago al seguro social obligatorio por el salario REAL PERCIBIDO DE GS. {datos_cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?']} Y POR TODO EL TIEMPO QUE DURE LA RELACION LABORAL, perjudicando de sobremanera al trabajador pues esta falta de aportes le perjudica en su futura jubilación y le causa un daño irreparable, pues como bien es sabido cientos de trabajadores PARAGUAYOS, son explotados y en su vejez no acceden a la jubilación por el incumplimiento de las normativas laborales. SE RECLAMA UNA INDEMNIZACIÓN DE GS. 500.000.000 (GUARANÍES QUINIENTOS MILLONES), POR EL DAÑO MORAL CAUSADO DENTRO DEL CONTRATO DE TRABAJO Y LA RELACIÓN LABORAL ANTE EL NO PAGO DE LA SEGURIDAD SOCIAL. '
    agregar_parrafo_con_negrita(doc,p12,['SE RECLAMA UNA INDEMNIZACIÓN DE GS. 500.000.000 (GUARANÍES QUINIENTOS MILLONES),'])
    
    
    doc.save(f'Generado/Demanda {datos_cliente['Nombres y Apellidos completos como esta en tu Cedula.'].split()[0]} contra {datos_cliente['Empresa en la que trabajo <Razon Social>']}.docx')
    