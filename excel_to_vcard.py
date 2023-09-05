import os
import pandas as pd
import vobject

def create_vcard(row, df):
    """
    Given a row from the dataframe, create a vCard object.
    :param row: Individual row from the dataframe containing user details.
    :param df: The entire dataframe (used for checking columns).
    :return: vobject.vCard object.
    """

    # Helper function to add values to the vCard if they exist.
    def add_value(card, field, value, **kwargs):
        if pd.notna(value):
            line = card.add(field)
            line.value = value
            for k, v in kwargs.items():
                setattr(line, k, v)
        return card

    # Initialize vCard and set name fields.
    card = vobject.vCard()
    card.add('n')
    card.n.value = vobject.vcard.Name(family=row['Last Name'], given=row['First Name'])
    card.add('fn')
    card.fn.value = f"{row['First Name']} {row['Last Name']}"

    # Add email, mobile phone, company, title, and company website if they exist.
    card = add_value(card, 'email', row['E-Mail'], type_param='INTERNET')
    card = add_value(card, 'tel', row['Mobile Phone'], type_param='CELL')
    card = add_value(card, 'title', row['Title'])
    card = add_value(card, 'url', row['Company Website'])

    # Add company and address if they exist.
    if pd.notna(row['Company']):
        card.add('org')
        card.org.value = [row['Company']]
    if pd.notna(row['Company Address']):
        card.add('adr')
        card.adr.value = vobject.vcard.Address(street=row['Company Address'])

    # Add LinkedIn profile if the column exists and the value for the row is not NA.
    if 'Linkedin Profile' in df.columns and pd.notna(row['Linkedin Profile']):
        linkedin_url = row['Linkedin Profile']
        if not linkedin_url.startswith("http"):
            linkedin_url = "https://" + linkedin_url
        card.add('x-socialprofile')
        card.x_socialprofile.value = linkedin_url
        card.x_socialprofile.type_param = 'linkedin'

    return card

def excel_to_vcards(filename):
    """
    Convert an Excel file into a list of vCard objects.
    :param filename: Path to the Excel file.
    :return: List of vobject.vCard objects.
    """
    df = pd.read_excel(filename)
    return [create_vcard(row, df) for _, row in df.iterrows()]

def save_vcards_separate(vcards, output_folder):
    """
    Save each vCard object as a separate .vcf file.
    :param vcards: List of vobject.vCard objects.
    :param output_folder: Directory where the .vcf files will be saved.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for card in vcards:
        # Convert special characters in names for filename compatibility.
        first_name = card.n.value.given.translate(str.maketrans('ıöüğçşİÖÜĞÇŞ', 'iougcsIOUGCS'))
        last_name = card.n.value.family.translate(str.maketrans('ıöüğçşİÖÜĞÇŞ', 'iougcsIOUGCS'))
        filename = f"{first_name}-{last_name}.vcf".lower()
        filepath = os.path.join(output_folder, filename)

        with open(filepath, 'w') as f:
            f.write(card.serialize())

if __name__ == '__main__':
    # Define the path to your Excel file and the output directory for the vCards.
    filename = '/path/to/your/excel/file.xlsx'
    output_folder = '/path/to/output/directory'

    vcards = excel_to_vcards(filename)
    save_vcards_separate(vcards, output_folder)
