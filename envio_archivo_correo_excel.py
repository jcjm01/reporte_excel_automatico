import pymysql
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#En este apartado se configura la conexion de la BD con python
conector = pymysql.connect(
    host="aqui_va_Tu_host",
    user="aqui_va_tu_user",
    password="aqui_va_tu_contraseña",
    database="aqui_va_tu_bd"
)
#############
try:
    #Aqui creamos un cursor para hacer las consultas a la BD
    cursor = conector.cursor()

    #Este query es un INNER JOIN que consultas tres tablas en mysql para generar el reporte
    query = " SELECT cliente.IdCliente,cliente.Departamento,maquina.Capacidad,maquina.RAM,maquina.SO,maquina.ESTADO,vcenters.IP,vcenters.ESTATUS FROM cliente INNER JOIN maquina ON cliente.IdCliente = maquina.IdCliente  INNER JOIN vcenters ON cliente.IdCliente= vcenters.IdCliente WHERE 'IdCliente'<= 10;"

    # Ejecuta la consulta en mysql
    cursor.execute(query)

    # Guarda los resultados de la query y los guarda en un DataFrame de la libreria de pandas
    resultados = cursor.fetchall()
    column_names = [i[0] for i in cursor.description]
    df = pd.DataFrame(resultados, columns=column_names)

    #Guarda el DataFrame anterior en un Excel
    archivo_excel = "resultados2.xlsx"
    df.to_excel(archivo_excel, index=False)

    #En este apartado se configura el apartado del envio de correo
    envia_email = "aqui_va_el_correo_de_quien_envia_el_correo"
    recibe_email = "aqui_va_el_correo_de_quien_recibe_el_correo"
    
    asunto = "Correo de prueba con archivo Excel adjunto"
    cuerpo_email = "Email de prueba de envio automatizado de correos que se ejecutauna vez al dia adjuntando un excel con un reporte generado de una BD en MYSQL."

    #Arma el mensaje del mail con los destinatarios y el que envia el correo
    msg = MIMEMultipart()
    msg["From"] = envia_email
    msg["To"] = recibe_email
    msg["Asunto"] = asunto 

    #Anexa el cuerpo del correo  
    msg.attach(MIMEText(cuerpo_email, "plain"))

    #Este bloque adjunta el archivo excel generado en la consulta INNER JOIN
    with open(archivo_excel, "rb") as anexar_excel:
        part = MIMEApplication(anexar_excel.read())
        part.add_header("Content-Disposition", f"attachment; filename= {archivo_excel}")
        msg.attach(part)

    #En este bloque se configura el servidor SMTP con puerto del mismo y el correo desde donde lo enviamos
    smtp_server = "smtp-mail.outlook.com"
    smtp_puerto = 587
    smtp_usuario = "aqui_va_tu_correo"
    smtp_passwd = "aqui_va_tu_la_contraseña_de_tu_correo"

    #Este bloque inicia el servicio SMTP
    servidorsmtp = smtplib.SMTP(smtp_server, smtp_puerto)
    servidorsmtp.starttls()

    # Este bloque inicia sesión del server SMTP
    servidorsmtp.login(smtp_usuario, smtp_passwd)

    #Aqui se envia el correo electornico
    servidorsmtp.sendmail(envia_email, recibe_email, msg.as_string())

    #Este bloque cierra la conexión con el SMTP
    servidorsmtp.quit()
#Se uso el metodo try-except para que en caso de fallar el programa no deje de correr pero muestre
#un mansaje de error
except Exception as e:
    print(f"Error: {e}")

finally:
    #Este modulo termina el cursor y la conexión
    cursor.close()
    conector.close()
########