import base64

if __name__ == '__main__':
    with open("email.ico", "rb") as i:
        b64str = base64.b64encode(i.read())
        img_code = '<img src="data:image/png;base64,' + b64str.decode('utf-8') + '" />'
    with open("img.html", "wb") as f:
        f.write(img_code.encode('utf-8'))
