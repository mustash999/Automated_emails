'''
 # @ ------------------------------------------------------------------------------: @ #

 # @ Author: 							Genius Panda								 @ #
 # @   : 																			:@ #
 # @ Create Time: 						23-10-08 21:22								 @ #
 # @ Modified time: 					24-11-17 18:35								 @ #
 # @ Modified by: 						Mustash										 @ #
 # @    : 																			:@ #
 # @ web: 								www.geniuspandatech.com						 @ #
 # @ Description: 						Put the description here
 # @ - ----------------------------------------------------------------------------: @ #

 '''


import		pandas as pd
import		os

original_excel= os.path.join( os.getcwd() ,'res', 'from_email.xlsx')	
company_keys= os.path.join( os.getcwd() ,'res', 'key.xlsx')
output_excel= os.path.join( os.getcwd() ,'res', 'Contacts.xlsx')

def prep_xcl_1(input_file, output_file):
	# Read the Excel file into a pandas DataFrame
	try:
		df = pd.read_excel(input_file)
	except Exception as e:
		print(f"Error reading the input file: {e}")
		return

	# Check if the DataFrame has at least 2 columns
	if df.shape[1] < 2:
		print("The Excel file should have at least 2 columns.")
		return

	# Extract the email addresses from the second column
	df['Email'] = df.iloc[:, 1].str.lower()  # Assuming the email column is the second column (0-indexed)

	# Extract domain names from email addresses
	df['Domain'] = df['Email'].str.split('@').str[1]

	# Drop duplicate rows based on the 'Email' column
	df = df.drop_duplicates(subset='Email')

	# Sort the DataFrame based on the 'Domain' column
	df = df.sort_values(by='Domain')

	# Remove the 'Domain' column if you don't want it in the output
	#df.drop('Domain', axis=1, inplace=True)

	# Save the sorted DataFrame to a new Excel file
	try:
		df.to_excel(output_file, index=False, engine='openpyxl', sheet_name='List')
		print(f"File '{output_file}' created with reordered rows (excluding duplicates).")
	except Exception as e:
		print(f"Error saving the output file: {e}")

########################################################################

def ms_name_ext(filename):
	# Read the Excel file
	df = pd.read_excel(filename, sheet_name='List')
	for index, row in df.iterrows():
		name = row['Name']
		if ',' in name:
			name_parts = name.split(',')
			last_name = name_parts[0].strip()
			first_name = name_parts[1].strip().split()[0]
		else:
			name_parts = name.split(' ')
			first_name = name_parts[0]
			if len (name_parts) > 1:
				last_name = name_parts[1]

		print(f"{name} is called {first_name} {last_name}")
		df.loc[index, 'First_Name'] = first_name
		df.loc[index, 'Last_Name'] = last_name
	print(df["First_Name"])
	df.to_excel(filename, sheet_name='List', index=False)
	  
########################################################################

def complete_company():
	df = pd.read_excel(output_excel, sheet_name='List')
	df2 = pd.read_excel(company_keys, sheet_name='Key')

	df['Domain'] = df['Domain'].str.lower()
	df2['Domain'] = df2['Domain'].str.lower()
	
	for index, row in df.iterrows():
		if row['Domain'] in df2['Domain'].values:
			comp_v = df2.loc[df2['Domain'] == row['Domain'], 'Company'].values[0]
			df.loc[index, 'Company'] = comp_v
			print(f"{index} {row['Email']} the company is {comp_v} ")

	df.to_excel(output_excel, sheet_name='List', index=False)	

########################################################################

def main():
	prep_xcl_1(original_excel, output_excel)
	ms_name_ext(output_excel)
	complete_company()

if __name__ == '__main__':
	main()