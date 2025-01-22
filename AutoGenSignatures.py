import pandas as pd
import os

# Load the Excel file
file_path = "employees_spsi.xlsx"  # Ensure this matches the actual file name
df_uploaded = pd.read_excel(file_path)

# Updated HTML template with structured address
html_template_with_address = """
<table
  style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #333; width: 100%; max-width: 600px; border-spacing: 0;">
  <tr>
    <td style="padding: 10px; vertical-align: top; text-align: left;">
      <p style="margin: 0; font-weight: bold; font-size: 16px;">{name}</p>
      <p style="margin: 0; font-size: 14px;">{title}</p>
      <p style="margin: 10px 0 0; font-size: 14px;">
        SPSI, Inc.<br>
        9825 85th Avenue N<br>
        Maple Grove MN 55369
      </p>
      <p style="margin: 10px 0 0; font-size: 14px;">
        <strong>Main:</strong> {phone_main} {extension_info} | <strong>Direct:</strong> {phone_direct}
      </p>
      <p style="margin-top: 10px 0; margin-bottom: 0;">
        <a href="https://www.spsi.com" style="color: #0078D4; text-decoration: none;">www.spsi.com</a>
      </p>
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding: 5px; text-align: left; border-top: 1px solid #ccc;">
      <img src="https://github.com/BOSSZEUS/1x/blob/master/50year-small.jpg?raw=true" alt="50 Years Anniversary"
        style="height: 125px; margin-left: 43px; margin-top: 5px;">
      <div style="margin: 10px;">
        <a href="https://x.com/i/flow/login?redirect_after_login=%2FSPSIINC"
          style="text-decoration: none; margin: 0 5px;">
          <img src="https://github.com/BOSSZEUS/1x/blob/master/X%20Logo.png?raw=true" alt="x"
            style="width: 24px; height: 24px;">
        </a>
        <a href="https://www.facebook.com/SPSIINC/?ref=ts&fref=ts" style="text-decoration: none; margin: 0 5px;">
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
          <img src="https://raw.githubusercontent.com/BOSSZEUS/1x/refs/heads/master/spsi-screenprinting_135x68.webp" alt="SPSI"
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

# Function to format the extension field
def format_extension(extension):
    # Safely handle extensions and convert to integer without decimals
    if pd.notna(extension) and extension != "":
        return f"<strong>Ext.</strong> {int(float(extension))}"  # Convert to integer
    return ""

# Generate HTML files
output_folder_with_address = "signatures_with_address/"
os.makedirs(output_folder_with_address, exist_ok=True)  # Ensure output directory exists

for _, row in df_uploaded.iterrows():
    extension_info = format_extension(row["Extension"])  # Add extension only if not blank
    html_content = html_template_with_address.format(
        name=row["Name"],
        title=row["Title"],
        phone_main=row["PhoneMain"],
        extension_info=extension_info,
        phone_direct=row["PhoneDirect"]
    )
    # Save the HTML file for each employee
    file_name = f"{row['Name'].replace(' ', '_')}_signature.html"
    with open(os.path.join(output_folder_with_address, file_name), "w") as f:
        f.write(html_content)

print(f"Signatures generated in {output_folder_with_address}")