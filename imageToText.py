from google.cloud import vision


def detect_text_uri(uri):
    """Detects text in the file located in Google Cloud Storage or on the Web.
    """
    from google.cloud import vision
    client = vision.ImageAnnotatorClient()
    image = vision.Image()
    image.source.image_uri = uri

    response = client.text_detection(image=image)
    texts = response.text_annotations
    print('Texts:')
    print(texts)
    # for text in texts:
    #     print(type(text.description))

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))


if __name__ == '__main__':
    detect_text_uri(
        "https://image.shutterstock.com/image-vector/sold-stamp-square-grunge-sign-260nw-1502617838.jpg")
