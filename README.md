# Saskaitu-valdymas

## Functionality:
1. Takes information from an Excel file composed of these columns in no particular order:
   - __Numeris__ - The number of the user
   - __Pirkejas__ - The name of the user
   - __Kaina__ - The price or amount
   - __Serija__ - The series or batch number
   - __Pavadinimas__ - The product name or description
   - __Gmail__ - The email address of the user

2. Column names can be incorrect to a certain extent

3. Writes the information to User objects

4. Uses the user objects to create PDF invoices.

5. The app then takes users gmail and the generated password inside of the "Prisijungti"

6. Sends PDF invoices to the corresponding email addresses.

## App
 - Made using python
 - In the App folder - Sąskaitų_valdymas. 
 - Can be made into a seperate app using information from the installing_app_cmd.txt file

## Custom password
If the user wishes to send these pdf files, he must generate an app password:

1. __2FA__ - turn on two factor authentification on their desired email account

2. Go to https://support.google.com/accounts/answer/185833?hl=en, scroll down and "press Create and manage your app passwords"

3. Sign in into your desired gmail account

4. Create the password, store it, and use it when prompted in the sign in part of the app

## Example
- App takes info from "Lentele CBC saskaitu"
- Generates invoice based on serie and name in the same folder as the app
- Now the user can send this invoice and a message to the person whose gmail was presented