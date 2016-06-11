import base64

class Binary2Image():
	'''This class will convert the binary data
						to an image'''

	
	def __init__(self, binary_data, image_name):
		self.binary_data = binary_data
		self.image_name = image_name

	def convert2image(self):
		image = base64.b64decode(self.binary_data)

		with open(self.image_name, 'w+') as img:
			img.write(image)