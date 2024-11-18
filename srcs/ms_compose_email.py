'''
 # @ ------------------------------------------------------------------------------: @ #

 # @ Author: 							Genius Panda								 @ #
 # @   : 																			:@ #
 # @ Create Time: 						23-10-01 21:10								 @ #
 # @ Modified time: 					24-10-31 22:42								 @ #
 # @ Modified by: 						Mustash										 @ #
 # @    : 																			:@ #
 # @ web: 								www.geniuspandatech.com						 @ #
 # @ Description: 						Put the description here
 # @ - ----------------------------------------------------------------------------: @ #

This script is part of the panpost project.
this is the final step befor sending the emails
this script will open the excel look for the language column and open the corresponding template
then it will replace the variables in the template with the values from the excel file
Author: mustash
last upuate: 17 Nov. 2024
'''       

import				pandas as pd
import				os
import				dotenv

dotenv.load_dotenv()
filename = os.path.join( os.getcwd() ,'res', 'Contacts.xlsx')
english_template = os.path.join( os.getcwd() ,'res', 'template', 'panpost_En_template.txt')
spanish_template = os.path.join( os.getcwd() ,'res', 'template', 'panpost_Sp_template.txt')
german_template = os.path.join( os.getcwd() ,'res', 'template', 'panpost_De_template.txt')
formal_german_template = os.path.join( os.getcwd() ,'res', 'template', 'panpost_Def_template.txt')


def ms_xclextract(file_name):
	try:
		# Read the Excel file into a DataFrame
		df = pd.read_excel(file_name)

		# Convert the DataFrame into a list of dictionaries (one dictionary per row)
		variables_list = df.to_dict(orient='records')

		return variables_list
	
	except FileNotFoundError:
		print(f"File '{file_name}' not found.")
	except Exception as e:
		print(f"An error occurred: {str(e)}")

###########################################################################

def ms_edittxt(txt_file, dict):
	try:
		# Open the file for reading
		with open(txt_file, 'r', encoding='utf-8' ) as file:
			content = file.read()

		# Replace variables in square brackets with the provided values
			for variable_name, variable_value in dict.items():
				content = content.replace(f'[{variable_name}]', str(variable_value))

		return content
	
	except FileNotFoundError:
		print(f"File '{txt_file}' not found.")
	except Exception as e:
		print(f"An error occurred: {str(e)}")

########################################################################

def ms_compose_email(filename):
	out_dir = 'msgs_to_send'
	variables_list = ms_xclextract(filename)
	for variables in variables_list:
		if variables['Language'] == 'D':
			edited_text = ms_edittxt(german_template, variables)
		elif variables['Language'] == 'E':
			edited_text = ms_edittxt(english_template, variables)
		elif variables['Language'] == 'S':
			edited_text = ms_edittxt(spanish_template, variables)
		elif variables['Language'] == 'Df':
			edited_text = ms_edittxt(formal_german_template, variables)
		# Save the edited text to a new file
		new_file_name = f"{variables['Email']}.txt"
		with open(os.path.join(out_dir, new_file_name), 'w', encoding='utf-8') as new_file:
			new_file.write(edited_text)
		print(f"File '{new_file_name}' created.")

###########################################################################


def main():
	ms_compose_email(filename)

if __name__ == "__main__":
	main()
