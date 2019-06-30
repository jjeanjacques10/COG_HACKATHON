import cognitive_face as CF

class ReconhecimentoMicrosoft():

    def indentificarFaceImagem(self, imgURL):
        KEY = '2dae3d599d09438fb0d33769b3abf948'
        location = 'westus'
        CF.Key.set(KEY)

        BASE_URL = 'https://{}.api.cognitive.microsoft.com/face/v1.0/detect?returnFaceId=true&returnFaceLandmarks=false&returnFaceAttributes=age,gender,smile,glasses,emotion,hair&recognitionModel=recognition_01&returnRecognitionModel=true'.format(
            location)  # Replace with your regional Base URL
        CF.BaseUrl.set(BASE_URL)

        result = CF.face.detect(imgURL)

        result = result[0]['faceAttributes']

        response = dict()
        print(result)
        if(result['smile'] >= 0.5):
            response["emocao"] = "Feliz"
        else:
            response["emocao"] = "Não Feliz"
        if (result['gender'] == 'male'):
            response["sexo"] = "Homem"
        elif (result['gender'] == 'female'):
            response["sexo"] = "Mulher"
        response["idade"] = result['age']
        print(response)
        return response
