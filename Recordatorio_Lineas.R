####################################################################-
##########  Correo recordatorio de requerimientos enviados  ########-
############################# By LE ################################-

################ I. Librerías, drivers y directorio ################

# I.1 Librerías y drivers

# i) Librerias 
#install.packages("dplyr")
library(dplyr)
#install.packages("lubridate")
library(lubridate)
#install.packages("readxl")
library(readxl)
#install.packages("stringr")
library(stringr)
#install.packages("purrr")
library(purrr)
#install.packages("blastula")
library(blastula)
#install.packages("kableExtra")
library(kableExtra)
#install.packages("googledrive")
library(googledrive)
#install.packages("googlesheets4")
library(googlesheets4)
#install.packages("httpuv")
library(httpuv)
#install.packages("tidyr")
library(tidyr)

# ii) Google
correo_usuario <- ""
drive_auth(email = correo_usuario) 
gs4_auth(token = drive_auth(email = correo_usuario), 
         email = correo_usuario)


#-----------------------------------------------------------------

################## II. Descarga de informacion ##################

# II.1 ConexiOn a base de datos
REQ_OSPA <- ""
REQ_SEFA <- ""
REQ_COFEMA1 <- ""
REQ_COFEMA2 <- ""
REQ_SINADA <- ""
BASE_SEFA <- ""

# i) Descarga y lectura de tablas
# Diccionario de datos
DIC_DATOS <- read_sheet(BASE_SEFA, sheet = "DIC_DATOS")
# OSPA
INSUMOS_OSPA <- read_sheet(REQ_OSPA, sheet = "SEGUIMIENTO", skip = 2)
# SEFA
INSUMOS_REQ_SEFA <- read_sheet(REQ_SEFA, sheet = "INSUMOS")
# COFEMA
INSUMOS_COFEMA1 <- read_sheet(ss=REQ_COFEMA1,sheet = "Registro")
INSUMOS_COFEMA1 <- INSUMOS_COFEMA1[ 1:829,]
INSUMOS_COFEMA2 <- read_sheet(ss=REQ_COFEMA2,sheet = "REQ LINEAS")
# SINADA
DERIVACIONES <- as.data.frame(read_sheet(REQ_SINADA, sheet = "Derivaciones"))

# II.2 Conexion a base de contactos
CONTACTOS <- ""
# Descarga de la tabla
CONTACTOS_OEFA <- read_sheet(CONTACTOS, sheet = "OEFA", skip = 1)

#-----------------------------------------------------------------

################## III. Procesamiento de datos ##################
# Diccionario de datos
# i) OSPA ----
DIC_DATOS_OSPA <- DIC_DATOS %>%
  filter(!is.na(NOMBRE_OSPA))

TABLA_OSPA <- INSUMOS_OSPA %>%
  filter(!is.na(`FECHA DE NOTIFICACION`))  %>%
  filter(`TIPO DE EFA` == "Directa" |
           `TIPO DE EFA` == "ODE/OE") %>%
  filter(TAREA == "Pedido de información" |
           TAREA == "Pedido de información adicional" |
           TAREA == "Pedido de información urgente") %>%
  filter(`ESTADO O ACTIVIDAD PENDIENTE` != "Terminado con reiterativo") %>%
  select(DIC_DATOS_OSPA$NOMBRE_OSPA)  %>%
  mutate(AREA_SEFA = "OSPA")

colnames(TABLA_OSPA) <- DIC_DATOS$CODIGO_SCRIPT

# ii) Requerimientos SEFA ----
DIC_DATOS_REQ_SEFA <- DIC_DATOS %>%
  filter(!is.na(NOMBRE_REQ_SEFA))

TABLA_REQ_SEFA <- INSUMOS_REQ_SEFA %>%
  filter(COD_OEFA!=" ") %>%
  select(DIC_DATOS_REQ_SEFA$NOMBRE_REQ_SEFA)  %>%
  mutate(AREA_SEFA = "SEFA")

colnames(TABLA_REQ_SEFA) <- DIC_DATOS$CODIGO_SCRIPT

# iii) SINADA ----
# Seleccion de denuncias
DERIVACIONES_1 <- DERIVACIONES %>%
  select(CODIGO = 'Código Sinada',
         HT,
         F_INGRESO  = 'Fecha de notificación del documento [Ver cargo]',
         N_DOC = 'Tipo y N° de documento',
         F_VENC = 'Fecha de vencimiento',
         LINEA = 'Destinatario',
         AREA = 'AREA_OEFA',
         FECHA_FIRMA = 'Fecha firma del documento',
         TIPO_PEDIDO = 'Acciones adoptadas',
         PLAZO = 'Plazo otorgado para la respuesta [Días hábiles]',
         RPTA_FINAL = 'Respuesta final',
         RPTAENPLAZO = 'Respondido dentro o fuera del plazo',
         SEGUIMIENTO2 = 'Seguimiento a la respuesta al documento emitido [Casilla automática]',
         F_RPTA = 'Fecha rpta final',
         EFA_ABRE = 'EFA ABREVIADO'
) %>%
  mutate(HT = case_when(as.character(HT) == "NULL" ~ "",
                        TRUE ~ as.character(HT)))

TABLA_SINADA <- DERIVACIONES_1 %>%
  filter(EFA_ABRE == 'OEFA',
         !is.na(F_VENC),
         !is.na(FECHA_FIRMA),
         !is.na(PLAZO),
         TIPO_PEDIDO != 'Informa registro (sin requerimiento)') %>%
  mutate(ESTADO = ifelse(RPTAENPLAZO == 'Respondido fuera de plazo', 'Atendido fuera de plazo',
                         ifelse(RPTAENPLAZO == 'Respondido dentro del plazo', 'Atendido en plazo',
                                ifelse(RPTAENPLAZO == 'Vencido', 'Pendiente vencido',
                                       ifelse(RPTAENPLAZO == 'En plazo', 'Pendiente en plazo','Falta'))))) %>%
  filter(ESTADO != 'Falta') %>%
  mutate(AREA_SEFA= 'SINADA') %>%
  select(CODIGO, HT, F_INGRESO, N_DOC, F_VENC, LINEA, AREA, ESTADO, F_RPTA, AREA_SEFA)


# iv) COFEMA ----
FECHAINICIO_COFEMA <- as.Date("31/12/2018", "%d/%m/%Y")

TABLA_COFEMA1 <- INSUMOS_COFEMA1 %>%
  filter(`COORDINACION`!= "Otros",
         !is.na(`FECHA DE VENCIMIENTO`),
         !is.na(`EXP. COFEMA`)) %>%
  select(`EXP. COFEMA`, `HOJA DE TRAMITE`,`FECHA DE SOLICITUD`,`SOLICITUD`, `FECHA DE VENCIMIENTO`,`COORDINACION`,`DIRECCION`, `ESTADO`,`FECHA DE RESPUESTA`) %>%
  filter(`FECHA DE SOLICITUD`>FECHAINICIO_COFEMA) %>%
  arrange(`FECHA DE VENCIMIENTO`)%>%
  filter(`ESTADO`=="Pendiente vencido"|
           `ESTADO`=="Pendiente en plazo"|
           `ESTADO`=="Atendido fuera de plazo"|
           `ESTADO`=="Atendido en plazo") %>%
  mutate(`AREA REMITENTE`= "COFEMA")

colnames(TABLA_COFEMA1) <- DIC_DATOS$CODIGO_SCRIPT

TABLA_COFEMA2 <- INSUMOS_COFEMA2 %>%
  filter(`COORDINACION`!= "Otros",
         !is.na(`FECHA DE VENCIMIENTO`),
         !is.na(`CASO COFEMA`)) %>%
  select(`CASO COFEMA`, `HT`,`FECHA DE SOLICITUD`,`SOLICITUD`, `FECHA DE VENCIMIENTO`,`COORDINACION`,`DIRECCION`, `ESTADO`,`FECHA DE RESPUESTA`) %>%
  arrange(`FECHA DE VENCIMIENTO`)%>%
  filter(`ESTADO`=="Pendiente vencido"|
           `ESTADO`=="Pendiente en plazo"|
           `ESTADO`=="Atendido fuera de plazo"|
           `ESTADO`=="Atendido en plazo") %>%
  mutate(`AREA REMITENTE`= "COFEMA")

colnames(TABLA_COFEMA2) <- DIC_DATOS$CODIGO_SCRIPT

TABLA_COFEMA <- rbind(TABLA_COFEMA1, TABLA_COFEMA2)

# v) Consolidado de datos y subida de informacion ----
PIN_SEFA <- rbind(TABLA_OSPA, TABLA_REQ_SEFA, 
                  TABLA_SINADA, TABLA_COFEMA)

# Base para subir informacion
PIN_SEFA_BD <- PIN_SEFA %>%
  arrange(F_INGRESO) %>%
  mutate(F_INGRESO = ymd(F_INGRESO),
         F_RPTA = ymd(F_RPTA),
         F_VENC = ymd(F_VENC))
colnames(PIN_SEFA_BD) <- DIC_DATOS$NOMBRE_VARIABLE
write_sheet(PIN_SEFA_BD, ss= BASE_SEFA, sheet = "1) REQ SEFA")

#-----------------------------------------------------------------

################### IV. Parámetros y funciones ####################

# IV.1 Parametros para envío de correos ----

# I) Destinatarios SEFA y CODE
sefa_actores <- c("")

sefa_desarrolladores <- c("")

code_actores <- c("")


# IV.1.1 Email: Cabecera ----
Arriba <- add_image(
  file = "https://i.imgur.com/Ig7gq5Z.png",
  width = 1000,
  align = c("right"))
Cabecera <- md(Arriba)

# IV.1.2 Email: Pie de pagina ----
Logo_Oefa <- add_image(
  file = "https://i.imgur.com/ImFWSQj.png",
  width = 280)
Pie_de_pagina <- blocks(
  md(Logo_Oefa),
  block_text(md("Av. Faustino Sánchez Carrión N.° 603, 607 y 615 - Jesús María"), align = c("center")),
  block_text(md("Teléfonos: 204-9900 Anexo 7154"), align = c("center")),
  block_text("www.oefa.gob.pe", align = c("center")),
  block_text(md("**Síguenos** en nuestras redes sociales"), align = c("center")),
  block_social_links(
    social_link(
      service = "Twitter",
      link = "https://twitter.com/OEFAperu",
      variant = "dark_gray"
    ),
    social_link(
      service = "Facebook",
      link = "https://www.facebook.com/oefa.peru",
      variant = "dark_gray"
    ),
    social_link(
      service = "Instagram",
      link = "https://www.instagram.com/somosoefa/",
      variant = "dark_gray"
    ),
    social_link(
      service = "LinkedIn",
      link = "https://www.linkedin.com/company/oefa/",
      variant = "dark_gray"
    ),
    social_link(
      service = "YouTube",
      link = "https://www.youtube.com/user/OEFAperu",
      variant = "dark_gray"
    )
  ),
  block_spacer(),
  block_text(md("Imprime este correo electrónico sólo si es necesario. Cuidar el ambiente es responsabilidad de todos."), align = c("center"))
)

# IV.1.3 Email: Asunto y Botones ----
# i) Asunto
mes_actual <- month(now(), label=TRUE, abbr = FALSE)
mes_actual <- str_to_lower(mes_actual)

# ii) Boton al tablero en la Web
cta_button <- add_cta_button(
  url = "",
  text = "Tablero SEFA"
)

# IV.2 Funciones ----

# i) Creacion de tabla

#OSPA
tabla <- function(efa){
  # Referencia a tabla en el ambiente global
  TABLA <- PIN_SEFA %>%
    filter(LINEA == efa) %>%
    filter(ESTADO == "Pendiente en plazo" |
             ESTADO == "Pendiente vencido")  %>%
    select(HT, AREA_SEFA, CODIGO, ESTADO, F_VENC) %>%
    arrange(F_VENC) %>%
    mutate(F_VENC = format(F_VENC, "%d/%m/%Y"))
  TABLA
}

# ii) Seleccion de destinatarios
destinar <- function(x,
                     # Consideramos valores predeterminados
                     y = NA, z = NA){
  
  # Definimos los destinatarios eliminando posibles NA
  dest <- c(x, y, z)
  dest_f <- dest[!is.na(dest)]
  dest_f
  
}

# iii) Envío de correo
correo_auto <- function(efa, nombre,
                        d1, d2,
                        cc1, cc2, cc3){
  # i) Destinatarios 
  # Principales
  dest <- destinar(d1, d2)
  Destinatarios <- c(dest)

  
  # En copia
  cc <- destinar(cc1, cc2, cc3)
  cc_f <- c(cc, sefa_actores)
  Destinatarios_cc <- c(cc_f)
  
  
  # En copia oculta
  bcc <- sefa_desarrolladores
  Destinatarios_bcc <- c(bcc)
  
  
  # ii) Asunto
  Asunto <- paste("SEFA | Recordatorio semanal del ", 
                  day(now())," de ", mes_actual, " de ", year(now()))
  
  # iii) Generacion de tabla recordatorio
  TABLA_RECORDATORIO <- tabla(efa)

  # Diseño de tabla
  tabla_formato <- TABLA_RECORDATORIO %>%
    kbl(col.names = c("Registro SIGED",
                      "Área de Sefa remitente",
                      "Código",
                      "Estado",
                      "Fecha máxima para la respuesta"),
        align = c("c", "c", "c")) %>%
    kable_styling(font_size = 12, bootstrap_options = "basic" , latex_options = "HOLD_position", full_width = FALSE) %>%
    column_spec(1 , latex_valign = "m", width = "3cm") %>%
    column_spec(2 , latex_valign = "m", width = "2cm") %>%
    column_spec(3 , latex_valign = "m", width = "3cm") %>%
    column_spec(4 , latex_valign = "m", width = "3cm") %>%
    column_spec(5 , latex_valign = "m", width = "3cm") %>%
    row_spec(0, bold = TRUE, color = "white", background = "gray", align = "c")
  
  # IV.2.1 Email: Cuerpo del mensaje ----

  # Negrita en nombre de la EFA
  efa_n = paste0("**",efa,"**")
  # Texto en el correo
  Cuerpo_del_mensaje <- blocks(
    md(c("
Estimado/a,", nombre, "   
       ",
         efa_n,
         "
La Subdirección de Seguimiento de Entidades de Fiscalización Ambiental (Sefa) remite, en calidad de recordatorio periódico, el listado de documentos pendientes de
atención al día de hoy por parte de su área, con el fin de que sean atendidos dentro de los plazos otorgados y evitar la emisión de reiterativos.

A continuación el detalle de los requerimientos:")),
    
    md(c(tabla_formato)),
    md("
 
 Para mayor información puede visitar el tablero de control haciendo click en el siguiente botón:"
    ),
    md(c(cta_button)),
    md("
 
 ***
 **Tener en cuenta:**
 - Si tiene dudas sobre algún pedido, comuníquese con el área remitente.
 - Si desea incluir otro(s) destinatario(s) al recordatorio semanal, por favor, escriba al correo proyectossefa@oefa.gob.pe.
 - Este correo electrónico ha sido generado de manera automática, gracias al uso de lenguaje de programación de alto nivel que es un proyecto impulsado desde la Subdirección.
 
      ")
  )
  
  # IV.2.2 Email: Composicion ----
  email <- compose_email(
    header = Cabecera,
    body = Cuerpo_del_mensaje, 
    footer = Pie_de_pagina,
    content_width = 1000
  )
  
  # IV.2.3 Email: Envío ----
  smtp_send(
    email,
    to = Destinatarios,
    from = c("SEFA" = ""),
    subject = Asunto,
    cc = Destinatarios_cc,
    bcc = Destinatarios_bcc,
    credentials = creds_key(id = ""),
    verbose = TRUE
  )
}


# iv) Funcion robustecida de descarga y carga de informaciÃ³n
R_correo_auto <- function(efa, nombre,
                          d1, d2,
                          cc1, cc2, cc3){
  out = tryCatch(correo_auto(efa, nombre,
                             d1, d2,
                             cc1, cc2, cc3),
                 error = function(e){
                   correo_auto(efa,  nombre,
                               d1, d2,
                               cc1, cc2, cc3) 
                 })
  return(out)
}


#-----------------------------------------------------------------

################# V. Preparacion de datos ########################

# V.1 Seleccion de universo de EFAs 
TOTAL_PENDIENTES <- PIN_SEFA  %>%
  filter(ESTADO == "Pendiente en plazo" |
           ESTADO == "Pendiente vencido") %>%
  mutate(VENCIDOS = case_when(ESTADO == "Pendiente vencido" ~ 1, TRUE ~ 0),
         PLAZO = case_when(ESTADO == "Pendiente en plazo" ~ 1, TRUE ~ 0)) %>%
  group_by(LINEA, AREA) %>%
  summarise(PEDIDOS_PLAZO = sum(PLAZO), PEDIDOS_VENCIDOS = sum(VENCIDOS),PEDIDOS =n()) 
 
         
         

# V.2 Seleccion de EFAs afiliadas
CONTACTOS_LINEAS <- CONTACTOS_OEFA %>%
  filter(!is.na(DESTINATARIO)) %>%
  mutate(LINEA = DESTINATARIO)  %>%
  select(LINEA,
         NOMBRE, CARGO, OFICINA,
         DESTINATARIO_1, DESTINATARIO_2,
         CC1, CC2, CC3)

# V.3 Seleccion de poblacion de EFAs afiliadas al servicio
EFAS_CORREO <- merge(CONTACTOS_LINEAS, TOTAL_PENDIENTES)
EFAS_CORREO <- EFAS_CORREO %>%
  arrange(-PEDIDOS)

EFAS_CORREO_OD <- EFAS_CORREO%>%
                  filter(AREA == "OD")

#-----------------------------------------------------------------

#################### VI. Envío de correos ########################
# VI.1 Envio de correos a EFA ----
pwalk(list(EFAS_CORREO$LINEA,
           EFAS_CORREO$NOMBRE,
           EFAS_CORREO$DESTINATARIO_1,
           EFAS_CORREO$DESTINATARIO_2,
           EFAS_CORREO$CC1,
           EFAS_CORREO$CC2,
           EFAS_CORREO$CC3),
      slowly(R_correo_auto, 
             rate_backoff(10, max_times = Inf)))

# VI.2 Envío de correo a SEFA ----

# Asunto de validacion
Asunto_val <- paste("PIN SEFA | Recordatorio de requerimientos a líneas enviados la semana del ", 
                day(now())," de ", mes_actual, " de ", year(now()))

# Tabla de EFAS que recibieron correo
EFAS_CORREO_VAL <- EFAS_CORREO %>%
  arrange(-PEDIDOS)  %>%
  mutate(DESTINATARIO_F = paste0(NOMBRE,
                                 ", ",
                                 CARGO)) %>%
  select(LINEA, DESTINATARIO_F, PEDIDOS_PLAZO, PEDIDOS_VENCIDOS) %>%
  relocate(LINEA, DESTINATARIO_F, PEDIDOS_PLAZO, PEDIDOS_VENCIDOS)

# Numero de correos en negrita
n_correos <- paste0("**",nrow(EFAS_CORREO),"**")

# Diseño de tabla
tabla_formato_val <- EFAS_CORREO_VAL %>%
  kbl(col.names = c("Área",
                    "Jefe/Coordinador de Área",
                    "Req. pendientes en plazo",
                    "Req. pendientes vencidos"),
      align = c("c", "c", "c")) %>%
  kable_styling(font_size = 12, bootstrap_options = "basic" , latex_options = "HOLD_position", full_width = FALSE) %>%
  column_spec(1 , latex_valign = "m", width = "10cm") %>%
  column_spec(2 , latex_valign = "m", width = "8cm") %>%
  column_spec(3 , latex_valign = "m", width = "2cm") %>%
  column_spec(4 , latex_valign = "m", width = "2cm") %>%
  row_spec(0, bold = TRUE, color = "white", background = "gray", align = "c")

# Cuerpo del mensaje
Cuerpo_del_mensaje_validacion <- blocks(
  md(c("
Buenos días, Dante Guerrero Barreto:   

**Subdirector de Seguimiento de Entidades de Fiscalización Ambiental**

¡Se han enviado exitosamente los ", n_correos, " correos recordatorios de esta semana!
Cabe destacar que estos correos se envían gracias a la información consignada en los registros de
información de Sefa.
Estos recordatorios hacen referencia a requerimientos de información que aún no tienen respuesta.

A continuación el detalle de las áreas que recibieron el correo recordatorio:")),
  
  md(c(tabla_formato_val)),
  md("
 
 Para mayor información puede visitar el tablero de control haciendo click en el siguiente botón:"
  ),
  md(c(cta_button))
  )


# VI.2.1 Email: Composicion ----
email_val <- compose_email(
  header = Cabecera,
  body = Cuerpo_del_mensaje_validacion, 
  footer = Pie_de_pagina,
  content_width = 1000
)

# VI.2.2 Email: Envío ----
smtp_send(
  email_val,
  to = sefa_actores,
  from = c("SEFA" = ""),
  subject = Asunto_val,
  cc = sefa_desarrolladores,
  credentials = creds_key(id = ""),
  verbose = TRUE
)
# VI.3 Envío de correo a CODE ----

# Tabla de EFAS que recibieron correo
EFAS_CORREO_VAL_OD <- EFAS_CORREO_OD %>%
  arrange(-PEDIDOS)  %>%
  mutate(DESTINATARIO_F = paste0(NOMBRE,
                                 ", ",
                                 CARGO)) %>%
  select(LINEA, DESTINATARIO_F, PEDIDOS_PLAZO, PEDIDOS_VENCIDOS) %>%
  relocate(LINEA, DESTINATARIO_F, PEDIDOS_PLAZO, PEDIDOS_VENCIDOS)

# Numero de correos en negrita
n_correos_OD <- paste0("**",nrow(EFAS_CORREO_OD),"**")

# Diseño de tabla
tabla_formato_val_OD <- EFAS_CORREO_VAL_OD %>%
  kbl(col.names = c("Oficina Desconcentrada",
                    "Jefe de Oficina Desconcentrada",
                    "Req. pendientes en plazo",
                    "Req. pendientes vencidos"),
      align = c("c", "c", "c")) %>%
  kable_styling(font_size = 12, bootstrap_options = "basic" , latex_options = "HOLD_position", full_width = FALSE) %>%
  column_spec(1 , latex_valign = "m", width = "10cm") %>%
  column_spec(2 , latex_valign = "m", width = "8cm") %>%
  column_spec(3 , latex_valign = "m", width = "2cm") %>%
  column_spec(4 , latex_valign = "m", width = "2cm") %>%
  row_spec(0, bold = TRUE, color = "white", background = "gray", align = "c")

# Cuerpo del mensaje
Cuerpo_del_mensaje_validacion_OD <- blocks(
  md(c("
Buenos días, Alicia Rodriguez:   

**Coordinadora de Oficinas Desconcentradas**

¡Se han enviado exitosamente los ", n_correos_OD, " correos recordatorios de esta semana!
Cabe destacar que estos correos se envían gracias a la información consignada en los registros de
información de Sefa.
Estos recordatorios hacen referencia a requerimientos de información que aún no tienen respuesta.

A continuación el detalle de las OD que recibieron el correo recordatorio:")),
  
  md(c(tabla_formato_val_OD)),
  md("
 
 Para mayor información puede visitar el tablero de control haciendo click en el siguiente botón:"
  ),
  md(c(cta_button))
)


# VI.3.1 Email: Composicion ----
email_val <- compose_email(
  header = Cabecera,
  body = Cuerpo_del_mensaje_validacion_OD, 
  footer = Pie_de_pagina,
  content_width = 1000
)

# VI.3.2 Email: Envío ----
smtp_send(
  email_val,
  to = code_actores,
  from = c("SEFA" = ""),
  subject = Asunto_val,
  cc = sefa_actores,
  bcc = sefa_desarrolladores,
  credentials = creds_key(id = ""),
  verbose = TRUE
)
