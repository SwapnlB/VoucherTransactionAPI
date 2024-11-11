# VoucherTransactionAPI
Process the Tally Daybook XML files and make a spreadsheet with the given format


Steps to run: 

1. Clone the repo
2. Ensure Postman, python, mentioned packages in requirements.txt are installed.
3. Open terminal and run ProcessVoucharTransaction.py to start the server.
4. Now open PostMan tool to test the API.
	- Enter the url - http://localhost:5000/upload/ 
	- Select method type as POST, In Body --> form-data --> Enter 'XMLfile' as keyName and from the dropdown beside it, select source 'File' and upload your 'input.xml' file
	- Now in the top right, from Send button dropdown, select 'Send and Download', this will send the request and download the Response.xlsx format file.
