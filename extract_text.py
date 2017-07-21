from docx import Document
import glob

'''
function that takes a directory of a word file and extracts the text from its paragraphs (NOT HEADER/FOOTER)

Parameters
----------
directory - String that has the full directory/path of the word document

Returns
-------
String with all of the extracted text

'''
def extractText(directory):
	
	print(directory)	
	# Document object from word document
	document = 	Document(directory)
	
	# will contain each paragraph's text
	fullText = []
	
	# go through first 50 paragraphs and extract text
	count = 0
	for p in document.paragraphs:
	
		if count == 50:
			break
		if p.text == '':
			continue
		fullText.append(p.text)
		count += 1
	
	# create one large string and return (each paragraph separated by a new line'
	return '\n'.join(fullText)

	
'''
function that writes the new .txt file with extracted data

Parameters
----------
text - String with all of the extracted text
original_directory - String with the original path of Word document (in the form './docs/*.cdocx'

Returns
-------
None

'''
def writeText(text, original_directory):
	
	# get file name (without the full path and the .docx extension
	file_name = original_directory.split('/')[-1].split('.')[0]
	
	# create text file
	new_file = open('./txt/' + file_name + '.txt', 'w')
	
	# write the given text
	new_file.write(text)
	
	# close file
	new_file.close()
	
'''
main method that is immediately run after executing this file
'''
if __name__ == '__main__':
	
	# collect all files that are .docx format
	files = glob.glob('./docs/*.docx')
	
	for file in files:
	
		writeText(extractText(file),file)
	