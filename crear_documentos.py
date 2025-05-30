from docxtpl import DocxTemplate, InlineImage
import re
import uuid
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_UNDERLINE
from docx.shared import Pt, Inches, Cm, Mm
import requests
from io import BytesIO
from datetime import datetime
import qrcode
import os
import shutil





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
    fecha_formateada = f"{dia} días del mes de {meses[mes_num]} del año {año}"
    
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



def limpiar_carpeta(ruta_carpeta, eliminar_subcarpetas=False):
    """
    Elimina todos los archivos de una carpeta local, con opción para eliminar subcarpetas.
    
    Args:
        ruta_carpeta (str): Ruta absoluta o relativa de la carpeta a limpiar
        eliminar_subcarpetas (bool): Si es True, también elimina subcarpetas y su contenido
    
    Returns:
        tuple: (archivos_eliminados, errores)
    """
    archivos_eliminados = 0
    errores = 0
    
    try:
        # Verificar si la carpeta existe
        if not os.path.exists(ruta_carpeta):
            raise FileNotFoundError(f"La carpeta no existe: {ruta_carpeta}")
        
        # Recorrer todos los elementos en la carpeta
        for nombre in os.listdir(ruta_carpeta):
            ruta_completa = os.path.join(ruta_carpeta, nombre)
            
            try:
                if os.path.isfile(ruta_completa):
                    os.remove(ruta_completa)
                    archivos_eliminados += 1
                elif os.path.isdir(ruta_completa) and eliminar_subcarpetas:
                    shutil.rmtree(ruta_completa)
                    archivos_eliminados += 1  # Contamos la carpeta como un elemento eliminado
            except Exception as e:
                print(f"Error al eliminar {ruta_completa}: {str(e)}")
                errores += 1
        
        return (archivos_eliminados, errores)
    
    except Exception as e:
        print(f"Error general: {str(e)}")
        return (0, 1)

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


def generar_qr_inline(doc, enlace, ancho_mm=30):
    """
    Genera un código QR desde un enlace y lo devuelve como InlineImage para docxtpl.
    
    Parámetros:
        doc: instancia de DocxTemplate
        enlace: texto o URL a codificar en el QR
        ancho_mm: ancho del QR en milímetros (default: 30)
    
    Retorna:
        InlineImage listo para usarse en el contexto del template
    """
    # Crear imagen QR
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
    qr.add_data(enlace)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")

    # Guardar QR como imagen temporal
    nombre_archivo = f"Img/qr_{uuid.uuid4().hex}.png"
    img.save(nombre_archivo)

    # Crear InlineImage
    return InlineImage(doc, nombre_archivo, width=Mm(ancho_mm))

# Extrae el ID de la URL de Google Drive (si está en formato "open?id=")
def convertir_url_google_drive(url):
    if "open?id=" in url:
        id_imagen = url.split("open?id=")[-1]
        return f"https://drive.google.com/uc?export=download&id={id_imagen}"
    return url  # Ya es directa


def generar_imagen_inline(doc,url):
    """
    Descarga una imagen desde una URL completa.
    Verifica que sea una imagen válida.
    Devuelve la ruta local del archivo guardado.
    """
    nombre_archivo = f"Img/imagen_{uuid.uuid4().hex}.jpg"

    # Descargar el contenido
    url_valida = convertir_url_google_drive(url)
    response = requests.get(url_valida)

    # Validar tipo MIME
    content_type = response.headers.get('Content-Type', '')
    if 'image' not in content_type:
        raise Exception(f"La URL no contiene una imagen válida. Content-Type: {content_type}")

    # Validar con PIL
    try:
        img = Image.open(BytesIO(response.content))
        img.verify()
    except Exception as e:
        raise Exception("El contenido descargado no es una imagen válida") from e

    # Guardar imagen localmente
    with open(nombre_archivo, 'wb') as f:
        f.write(response.content)

    return InlineImage(doc,nombre_archivo)


def  FORM_DATOS_NUEVOS_PARA_TRABAJADOR(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'],
        'estado_civil': cliente['Estado Civil como esta en tu cedula'],
        'nacionalidad': cliente['Nacionalidad'],
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'],
        'barrio': cliente['Barrio'],
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'telefono': cliente['Telefono de contacto personal'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'],
        'direccion_empresa': cliente['Direccion de la Empresa'],
        'ruc_empresa':cliente['Ruc de la empresa'],
        'fecha_ingreso': cliente['Fecha de ingreso'],
        'fecha_despido': cliente['Fecha de Despido'],
        'jornada_laboral': cliente['JORNADA LABORAL. Como es o era tu Jornada Laboral? Lunes a Viernes, Lunes a Lunes?'],
        'horario_laboral': cliente['HORARIO DE TRABAJO. Como era tu Horario que Cumplias? Ej 8.00 a 18.00'],
        'salario': cliente['CUANTO ERA SALARIO. Mensual, semanal, diario?'],
        'ips': cliente['IPS'],
        'bonificacion_familiar':cliente['Bonficacion familiar por Cuantos hijos si la respuesta fue SI'],
        'tarea_realizada': cliente['Describe las tareas o funciones que desempenabas en el lugar de trabajo.'],
        'motivo_despido': cliente['MOTIVO DE DESPIDO. Cuentanos como se dio la situacion.'],
        'quien_lo_despidio': cliente['Quien te Comunico de tu despido?. Nombre Apellido, cargo en la empresa.'],
        'medio_del_despido': cliente['DESPIDO COMUNICACION. Como te comunicaron tu despido. Verbal, por escrito con nota, por llamada telefonica, por mensaje de texto?'],
        'salarios_pendientes': cliente['SALARIOS PENDIENTES. Te deben Salarios, Cuanto de cuantos dias o meses?'],
        'medio_pago_salario': cliente['PAGOS DE SALARIOS. Como recibias los pagos de salario o jornales?. Efectivo, Trasferencia, giros? via que banco?'],
        'vacaciones_pendientes': cliente['VACACIONES. Salias o tenias vacaciones? Te deben vacaciones?'],
        'aguinaldo_pendiente': cliente['AGUINALDO. Recibias Aguinaldo?. Te pagaban o te debe?'],
        'contaba_contrato': cliente['CONTRATO DE TRABAJO. Tenias contrato de Trabajo Firmado'],
        'ofrecio_liquidacion': cliente['LIQUIDACION.Te presentaron tu liquidacion de salarios y haberes al momento del despido?. Adjuntar Foto.'],
        'firmo_documento_en_blanco': cliente['Firmaste en algun momento algun Documento en blanco o pagare?'],
        'observaciones': cliente['Alguna informacion adicional que deseas agregar?'],
        'entrevistador':cliente['Entrevista realizada por']
    }

    fecha_actual = datetime.now().strftime('%d/%m/%Y')

    doc = DocxTemplate('Plantilla/Plantillla_Formulario_Datos.docx')


    imagen_actor = generar_imagen_inline(doc,'https://drive.google.com/open?id=1rO9qgzKcUCeyekHE16t8eFbxxU7V_IMq')
    link_ubicacion_actor = 'https://maps.app.goo.gl/gfjAkq5uewLpBV9QA'
    qr_actor = generar_qr_inline(doc,'https://maps.app.goo.gl/gfjAkq5uewLpBV9QA')
    imagen_demandado = generar_imagen_inline(doc,'https://drive.google.com/open?id=1lFkjyZLz4nmhXDDlmjGuaNdlQDKQQEAN')
    link_ubicacion_demandado = 'https://maps.app.goo.gl/c77RjPCYUGQr3Ti17'
    qr_demandado = generar_qr_inline(doc,'https://maps.app.goo.gl/c77RjPCYUGQr3Ti17')
    imagenesYfecha = {'fecha_hoy': fecha_actual,
                    'imagen_actor':imagen_actor,
                    'link_ubicacion_actor': link_ubicacion_actor,
                    'imagen_demandado':imagen_demandado,
                    'link_ubicacion_demandado':link_ubicacion_demandado,
                    'qr_actor':qr_actor,
                    'qr_demandado':qr_demandado}
    contexto.update(imagenesYfecha)


    doc.render(contexto)

    limpiar_carpeta('Img')

    doc.save(f'Generado/Formulario_Datos_{contexto['ci']}.docx')



def  Carta_Poder(cliente):
    contexto = {
        'fecha': obtener_fecha_formateada(),
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'estado_civil': cliente['Estado Civil como esta en tu cedula'].lower(),
        'nacionalidad': cliente['Nacionalidad'].lower(),
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'].capitalize(),
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'direccion_empresa': cliente['Direccion de la Empresa'],
        'ciudad_empresa': cliente['Ciudad de la empresa'].capitalize(),
        'ruc_empresa':cliente['Ruc de la empresa']
    }

    if cliente['Sexo']=='Femenino':
        contexto['estado_civil']=contexto['estado_civil'][:-1]+'a'


    doc = DocxTemplate('Plantilla/Carta_de_Poder.docx')



    doc.render(contexto)


    doc.save(f'Generado/Carta_de_Poder_{contexto['ci']}.docx')

def Carta_Compromiso(cliente):
   
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'fecha': obtener_fecha_formateada(),
        'ci': cliente['Numero de Cedula'],
        'ciudad': cliente['Ciudad'].capitalize(),
        'direccion_calle': cliente['Direccion Particular, Calles, Numero de casa'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'ruc_empresa':cliente['Ruc de la empresa']
    }

    doc = DocxTemplate('Plantilla/Carta_Compromiso.docx')

    doc.render(contexto)


    doc.save(f'Generado/Carta_Compromiso_{contexto['ci']}.docx')
    
    
def Desistimiento_de_renuncia(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'ci': cliente['Numero de Cedula']
    }

    doc = DocxTemplate('Plantilla/Desistimiento_de_Renuncia.docx')

    doc.render(contexto)


    doc.save(f'Generado/Desistimiento_de_renuncia_{contexto['ci']}.docx')

def Nota_de_Renuncia(cliente):
    contexto = {
        'nombre_completo': cliente['Nombres y Apellidos completos como esta en tu Cedula.'].upper(),
        'ci': cliente['Numero de Cedula'],
        'empresa_que_trabajo': cliente['Empresa en la que trabajo <Razon Social>'].upper(),
        'ruc_empresa':cliente['Ruc de la empresa']
    }

    doc = DocxTemplate('Plantilla/Nota_de_Renuncia.docx')

    doc.render(contexto)


    doc.save(f'Generado/Nota_de_Renuncia_{contexto['ci']}.docx')
    
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
    