# mass_email_sender

Explanation:

1. **Email Validation (`email_check`):**
   - The `email_check` function takes an email address as input and uses a regular expression to check if it's a valid email format.
   - If the email is valid, it calls the `send_mail` function; otherwise, it prints an error message.

2. **Email Sending (`send_mail`):**
   - The `send_mail` function creates an `EmailMessage` object, sets the sender, recipient, subject, and content.
   - It then establishes a connection to the Gmail SMTP server, logs in, sends the email, and prints a success message.

3. **Main Function (`main`):**
   - Loads an Excel file named 'example.xlsx' using the `openpyxl` library.
   - Prompts the user to enter the row and column numbers where email addresses are located.
   - Validates the user input to ensure it's a number greater than zero.
   - Iterates through the rows, retrieves email addresses, and calls the `email_check` function.

4. **Script Execution:**
   - When the script is executed (`if __name__ == '__main__':`), it calls the `main` function with the argument 'example.xlsx'.

Example:
If you want to provide the Excel file as a command-line argument when running the script in a bash environment, you can modify the last line of the script. Here's an example of how you can run the script with a command-line argument in bash:

```bash
python script_name.py example.xlsx
```

Replace `script_name.py` with the actual name of your Python script.

For example, if your script is named `email_sender.py`, you would run:

```bash
python email_sender.py example.xlsx
```

This assumes that the Excel file (`example.xlsx` in this case) is in the same directory as the script. If it's in a different directory, provide the full path or navigate to that directory in the terminal before running the script.

Make sure you have Python installed on your system and that it's in your system's PATH.

Please note:
- Replace 'your_email@gmail.com' and 'your_password_or_app_password' with your Gmail credentials. It's recommended to use an app-specific password for security.
- The script will send emails to the addresses specified in the Excel file, starting from the specified row and column.
- The script currently exits after processing one row due to the `sys.exit()` call. You may want to modify this behavior based on your requirements.