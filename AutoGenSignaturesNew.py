import pandas as pd
import os

# File path to the Excel file
file_path = r"C:\Users\jscheftic\OneDrive - spsi.com\Desktop\SPSI Code\SPSI Signatures\employees_spsi.xlsx"

# Load the Excel file with the correct engine
try:
    df_uploaded = pd.read_excel(file_path, engine='openpyxl')
    print("Excel file loaded successfully!")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()  # Exit the script if the file can't be loaded

# Updated HTML template with structured address
html_template_with_address = """
<table
  style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #333; width: 100%; max-width: 600px; border-spacing: 0;">
  <tr>
    <td style="padding: 10px; vertical-align: top; text-align: left;">
      <p style="margin: 0; font-weight: bold; font-size: 16px;">{name}</p>
      {title_section}
      <p style="margin: 10px 0 0; font-size: 14px;">
        SPSI, Inc.<br>
        9825 85th Avenue N<br>
        Maple Grove MN 55369
      </p>
      <p style="margin: 10px 0 0; font-size: 14px;">
        <strong>Main:</strong> {phone_main} {extension_info} {direct_section}
      </p>
      {email_section}
      <p style="margin-top: 10px 0; margin-bottom: 0;">
        <a href="https://www.spsi.com" style="color: #0078D4; text-decoration: none;">www.spsi.com</a>
      </p>
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding: 5px; text-align: left; border-top: 1px solid #ccc;">
      <!-- Added Celebration Text -->
      <div style="margin: 10px 0; font-size: 14px; color: #333;">
        <p style="margin: 0; font-weight: bold; font-size: 16px;">Celebrate 50 Years With Us</p>
        <p style="margin: 5px 0;">Join us September 23rd–25th for three days of seminars and a special anniversary celebration.</p>
        <p style="margin: 5px 0;">
          <a href="https://spsi.com/pages/50th-anniversary-celebration" style="color: #0078D4; text-decoration: underline;">
            Register here
          </a>
        </p>
      </div>

      <!-- Anniversary Image -->
      <img src="https://github.com/BOSSZEUS/1x/blob/master/50year-small.jpg?raw=true" alt="50 Years Anniversary"
        style="height: 125px; margin-left: 43px; margin-top: 5px;">

      <div style="margin: 10px;">
        <a href="https://x.com/i/flow/login?redirect_after_login=%2FSPSIINC"
          style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/X%20Logo.png?raw=true" alt="x"
            style="width: 24px; height: 24px;">
        </a>
        <a href="https://www.facebook.com/SPSIINC/?ref=ts&fref=ts"
          style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/FacebookLogo.png?raw=true" alt="Facebook"
            style="width: 24px; height: 24px;">
        </a>
        <a href="https://www.youtube.com/user/spsivideo" style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/youtube.png?raw=true" alt="YouTube"
            style="width: 24px; height: 24px;">
        </a>
        <a href="https://www.linkedin.com/company/spsi-incorporated/posts/?feedView=all" style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/LI-In-Bug.png?raw=true" alt="LinkedIn"
            style="width: 28px; height: 24px;">
        </a>
        <a href="https://spsi.com/blogs/events" style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/spsi-screenprintingPNG.png?raw=true" alt="SPSI"
            style="width: 40px; height: 24px;">
        </a>
      </div>
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding: 0px; font-size: 12px; color: #1f497d; text-align: left;">
      <p style="margin: 0;">Confidentiality Notice: All information pertaining to this email contains confidential
        information intended only for the use of the recipient(s) named in the header text. If you are not the intended
        recipient, you are hereby notified that any disclosure, copying, distribution, or the taking of any action in
        reliance on the contents of this emailed information except its direct delivery to the person named above is
        strictly prohibited. If you have received this email in error, please notify us immediately by replying to this
        email and delete all copies of this message. This message is protected by applicable legal privileges and is
        confidential.</p>
    </td>
  </tr>
</table>
"""

# Functions to handle conditional sections
def format_extension(extension):
    if pd.notna(extension) and extension != "":
        return f"<strong>Ext.</strong> {int(float(extension))}"  # Convert to integer
    return ""

def format_title(title):
    if pd.notna(title) and title != "":
        return f'<p style="margin: 0; font-size: 14px;">{title}</p>'
    return ""

def format_direct(phone_direct):
    if pd.notna(phone_direct) and phone_direct != "":
        return f'| <strong>Direct:</strong> {phone_direct}'
    return ""

def format_email(email):
    if pd.notna(email) and email != "":
        return f'<p style="margin: 0; font-size: 14px;"><strong>Email:</strong> <a href="mailto:{email}" style="color: #0078D4; text-decoration: none;">{email}</a></p>'
    return ""

# Generate HTML files
output_folder_with_address = "signatures_New/"
os.makedirs(output_folder_with_address, exist_ok=True)  # Ensure output directory exists

for _, row in df_uploaded.iterrows():
    extension_info = format_extension(row["Extension"])  # Add extension only if not blank
    title_section = format_title(row["Title"])  # Add title only if not blank
    direct_section = format_direct(row["PhoneDirect"])  # Add direct number only if not blank
    email_section = format_email(row["Email"])  # Add email only if not blank

    html_content = html_template_with_address.format(
        name=row["Name"],
        title_section=title_section,
        phone_main=row["PhoneMain"],
        extension_info=extension_info,
        direct_section=direct_section,
        email_section=email_section
    )
    # Save the HTML file for each employee
    file_name = f"{row['Name'].replace(' ', '_')}_signature.html"
    with open(os.path.join(output_folder_with_address, file_name), "w") as f:
        f.write(html_content)

print(f"Signatures generated in {output_folder_with_address}")