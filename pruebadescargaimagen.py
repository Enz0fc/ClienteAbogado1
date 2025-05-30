import requests


url = 'https://drive.google.com/open?id=1yU2nCTZ06-e2lu5rqbtbUucZPseObGMj'
image_path = 'imagen_temp.jpg'

response = requests.get(url)
if response.status_code == 200:
    with open(image_path, 'wb') as f:
        f.write(response.content)
else:
    raise Exception("No se pudo descargar la imagen.")