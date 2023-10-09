# AOJ

## Instalación

1. Ingresar al servidor sts2107 por escritorio remoto usando las credenciales del Credicoop (Solicitar permisos para poder operar con AOJ, hablar con Betty o referirse al Jira TEST-3543) 

2. Dentro del servidor ir a la carpeta "C:\temp\oj", copiar la carpeta y pegarla en la máquina del credi 

3. Una vez copiada, cerrar sesion del servidor 

4. Dentro de la carpeta "oj" se encuentra el "InstaladorAOJ7.exe", instalar la app 

5. Una vez instalada, copiar el archivo "SGB.ini" de la carpeta "oj" dentro de la carpeta donde se instaló la app en nuestra pc 

6. Abrir AOJ y verificar que no se muestre ningún error 

## Python

Para evitar el proxy y poder instalar packages:

- set http_proxy=http://username:password@proxyAddress:port 
- set https_proxy=https://username:password@proxyAddress:port

****IMPORTANTE:**** Ejecutar CMD como administrador

****Packages necesarios:****
- opencv-python
- Pillow
- PyAutoGUI
- pytesseract
- Pywin32
- pywinauto
- selenium
- Jpype1

****Tesseract:****
Es necesario instalar Tesseract para el reconocimiento de imagenes. El instalador se encuentra en: \\SFS-1\Testing\Automatizacion_de_Proyectos\Herramientas\tesseract

## Tips

| Tipo de apertura   | r |  r+ | w  | w+ | a | a+
|:---|:---:|:---:|:---:|:---:|:---:|:---:|
read              | + |  +  |    | +  |   | +
write             |   |  +  | +  | +  | + | +
write after seek  |   |  +  | +  | +  |   |
create            |   |     | +  | +  | + | +
truncate          |   |     | +  | +  |   |
position at start | + |  +  | +  | +  |   |
position at end   |   |     |    |    | + | +


