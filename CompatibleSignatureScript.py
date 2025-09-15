import pandas as pd
import os

# File path to the Excel file
file_path = r"C:\Users\jscheftic\SPSI Code\SPSI Signatures\employees_spsi.xlsx"

# Load the Excel file
try:
    df_uploaded = pd.read_excel(file_path, engine='openpyxl')
    print("Excel file loaded successfully!")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Outlook-safe HTML signature template with clean hyphen and dynamic address
html_template_with_address = """
<table style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #333; width: 100%; max-width: 600px; border-spacing: 0;" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding: 10px; vertical-align: top; text-align: left;">
      <p style="margin: 0; font-weight: bold; font-size: 16px;">{name}</p>
      {title_section}
      <p style="margin: 10px 0 0; font-size: 14px;">
        SPSI, Inc.<br>
        {address1}<br>
        {address2}
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

  <!-- Celebration Text Inserted -->
  <tr>
    <td colspan="2" style="padding: 5px 0 0 0; text-align: left; border-top: 1px solid #ccc;">
      <div style="margin: 10px 0; font-size: 14px; color: #333;">
        <p style="margin: 0; font-weight: bold; font-size: 16px;">Celebrate 50 Years With Us</p>
        <p style="margin: 5px 0;">Join us September 23rd-25th for three days of seminars and a special anniversary celebration.</p>
        <p style="margin: 5px 0;">
          <a href="https://spsi.com/pages/50th-anniversary-celebration" style="color: #0078D4; text-decoration: underline;">
            Register here
          </a>
        </p>
      </div>
    </td>
  </tr>

  <!-- Anniversary Image -->
  <tr>
    <td colspan="2" style="padding: 5px; text-align: left;">
      <!--[if gte mso 9]>
      <v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:125px;height:125px;">
        <v:imagedata src="https://github.com/BOSSZEUS/1x/blob/master/50year-small.jpg?raw=true" o:href="https://www.spsi.com" />
      </v:rect>
      <![endif]-->
      <![if !mso]>
      <img src="https://github.com/BOSSZEUS/1x/blob/master/50year-small.jpg?raw=true"
           alt="50 Years Anniversary"
           width="125" height="125"
           style="width: 125px; height: 125px; display: block;" />
      <![endif]>
    </td>
  </tr>

  <!-- Social Icons Row -->
  <tr>
    <td style="padding: 10px 0;">
      <table cellpadding="0" cellspacing="0" border="0" style="text-align: left;">
        <tr>
          <td style="padding: 0 5px;">
            <a href="https://x.com/i/flow/login?redirect_after_login=%2FSPSIINC">
              <img src="https://github.com/BOSSZEUS/1x/blob/master/X%20Logo.png?raw=true" alt="X" width="24" height="24" style="display: block; border: 0;" />
            </a>
          </td>
          <td style="padding: 0 5px;">
            <a href="https://www.facebook.com/SPSIINC/?ref=ts&fref=ts">
              <img src="https://github.com/BOSSZEUS/1x/blob/master/FacebookLogo.png?raw=true" alt="Facebook" width="24" height="24" style="display: block; border: 0;" />
            </a>
          </td>
          <td style="padding: 0 5px;">
            <a href="https://www.youtube.com/user/spsivideo">
              <img src="https://github.com/BOSSZEUS/1x/blob/master/youtube.png?raw=true" alt="YouTube" width="24" height="24" style="display: block; border: 0;" />
            </a>
          </td>
          <td style="padding: 0 5px;">
            <a href="https://www.linkedin.com/company/spsi-incorporated/posts/?feedView=all">
              <img src="https://github.com/BOSSZEUS/1x/blob/master/LI-In-Bug.png?raw=true" alt="LinkedIn" width="28" height="24" style="display: block; border: 0;" />
            </a>
          </td>
          <td style="padding: 0 5px;">
            <a href="https://spsi.com/blogs/events">
              <img src="https://github.com/BOSSZEUS/1x/blob/master/spsi-screenprintingPNG.png?raw=true" alt="SPSI" width="40" height="24" style="display: block; border: 0;" />
            </a>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- Footer -->
  <tr>
    <td style="font-size: 12px; color: #1f497d; text-align: left;">
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

# Utility formatters
def format_extension(extension):
    if pd.notna(extension) and extension != "":
        return f"<strong>Ext.</strong> {int(float(extension))}"
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

# Output directory
output_folder_with_address = "CompatibleSignatures"
os.makedirs(output_folder_with_address, exist_ok=True)

# Generate HTML signatures
for _, row in df_uploaded.iterrows():
    extension_info = format_extension(row["Extension"])
    title_section = format_title(row["Title"])
    direct_section = format_direct(row["PhoneDirect"])
    email_section = format_email(row["Email"])

    html_content = html_template_with_address.format(
        name=row["Name"],
        title_section=title_section,
        phone_main=row["PhoneMain"],
        extension_info=extension_info,
        direct_section=direct_section,
        email_section=email_section,
        address1=row["Address1"],
        address2=row["Address2"]
    )

    file_name = f"{row['Name'].replace(' ', '_')}_signature.html"
    with open(os.path.join(output_folder_with_address, file_name), "w", encoding="utf-8") as f:
        f.write(html_content)

print(f"Signatures generated in {output_folder_with_address}")
