import pypdf2
import subprocess

# Set the path to the John the Ripper executable
john_path = '/path/to/john'

# Open the PDF file in read-binary mode
with open('password-protected.pdf', 'rb') as file:
    # Create a PDF object
    pdf = pypdf2.PdfFileReader(file)

    # Check if the PDF is encrypted
    if pdf.isEncrypted:
        # Set the path to a wordlist containing possible passwords
        wordlist_path = '/path/to/wordlist'

        # Set the path to the file containing the hash of the PDF's password
        hash_path = '/path/to/hash_file'

        # Run John the Ripper with the given parameters
        result = subprocess.run([john_path, '--wordlist=' + wordlist_path, hash_path], stdout=subprocess.PIPE)

        # Get the cracked password from the output of John the Ripper
        cracked_password = result.stdout.decode().strip().split(' ')[-1]

        # Decrypt the PDF with the cracked password
        pdf.decrypt(cracked_password)

    # Iterate through all the embedded files in the PDF
    for index in range(pdf.getNumEmbeddedFiles()):
        # Get the embedded file object
        embedded_file = pdf.getEmbeddedFile(index)

        # Check if the embedded file is an Excel file
        if embedded_file.name.endswith('.xlsx'):
            # Write the Excel file to the current directory
            with open(embedded_file.name, 'wb') as output:
                output.write(embedded_file.content)
