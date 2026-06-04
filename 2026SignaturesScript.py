import pandas as pd
import os
import re

# File path to the Excel file
file_path = r"C:\Users\jscheftic\SPSI Code\SPSI Signatures\employees_spsi.xlsx"

# Load the Excel file
try:
    df_uploaded = pd.read_excel(file_path, engine='openpyxl')
    print("Excel file loaded successfully!")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# 2026 design signature template.
# Layout: header block (name/title/Mobile/Main+Ext/Email) with bottom border,
# mission strap, logo + address + website, right-aligned social icons,
# 5-segment color bar + confidentiality (combined in one TD to prevent
# Outlook from inflating spacing around the color bar on forward/reply).
html_template = """
<table style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #333; width: 100%; max-width: 600px; border-spacing: 0;" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding: 10px; vertical-align: top; text-align: left; border-bottom: 2px solid #333333;">
      <p style="margin: 0; font-weight: bold; font-size: 20px; color: #000000;">{name}</p>
      {title_section}
      {mobile_section}
      <p style="margin: 6px 0 0; font-size: 13px; color: #000000;">
        <strong>Main:</strong> <a href="tel:{phone_main_tel}" style="color: #000000; text-decoration: none;">{phone_main}</a> {extension_info} {direct_section}
      </p>
      {email_section}
    </td>
  </tr>

  <!-- Mission strap line -->
  <tr>
    <td style="padding: 8px 10px 0 10px; text-align: left;">
      <p style="margin: 0; font-size: 10px; color: #808080; letter-spacing: 0.5px;">DRIVEN TO SUCCEED &nbsp;|&nbsp; TO SERVE AMAZINGLY &nbsp;|&nbsp; TRUE TO OURSELVES &nbsp;|&nbsp; WILLINGNESS TO INVEST</p>
    </td>
  </tr>

  <!-- Logo + address -->
  <tr>
    <td style="padding: 12px 10px 0 10px; text-align: left;">
      <p style="margin: 0; font-size: 0; line-height: 0; mso-line-height-rule: exactly;">
        <a href="https://www.spsi.com" style="text-decoration: none; font-size: 0; line-height: 0;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="SPSI" width="130" height="37" style="width: 130px; height: 37px; display: block; border: 0; vertical-align: top;" /></a>
      </p>
      <p style="margin: 8px 0 0; font-size: 13px;">{address1}</p>
      <p style="margin: 0; font-size: 13px;">{address2}</p>
      <p style="margin: 0; font-size: 13px;">
        <a href="https://www.spsi.com" style="color: #333333; text-decoration: none;">www.spsi.com</a>
      </p>
    </td>
  </tr>

  <!-- Social Icons Row (right-aligned) -->
  <tr>
    <td style="padding: 0 10px; text-align: right;">
      <table cellpadding="0" cellspacing="0" border="0" align="right" style="border-spacing: 0;">
        <tr>
          <td style="padding: 0 3px;">
            <a href="https://veloxforum.com"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-forum.png?raw=true" alt="SPSI Forum" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
          <td style="padding: 0 3px;">
            <a href="https://www.linkedin.com/company/spsi-incorporated/posts/?feedView=all"><img src="https://allenbenedikt.github.io/spsi-email-assets/linkedin.png?raw=true" alt="LinkedIn" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
          <td style="padding: 0 3px;">
            <a href="https://www.instagram.com/spsiinc/"><img src="https://allenbenedikt.github.io/spsi-email-assets/instagram.png?raw=true" alt="Instagram" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
          <td style="padding: 0 3px;">
            <a href="https://www.facebook.com/SPSIINC/?ref=ts&fref=ts"><img src="https://allenbenedikt.github.io/spsi-email-assets/meta.png?raw=true" alt="Facebook" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
          <td style="padding: 0 3px;">
            <a href="https://x.com/i/flow/login?redirect_after_login=%2FSPSIINC"><img src="https://allenbenedikt.github.io/spsi-email-assets/x.png?raw=true" alt="X" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
          <td style="padding: 0 3px;">
            <a href="https://www.youtube.com/user/spsivideo"><img src="https://allenbenedikt.github.io/spsi-email-assets/youtube.png?raw=true" alt="YouTube" width="34" height="34" style="width: 34px; height: 34px; display: block; border: 0;" /></a>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- Color bar + Confidentiality Footer (combined to avoid extra TR boundary around the color bar) -->
  <tr>
    <td style="padding: 10px; font-size: 9px; color: #333333; text-align: left;">
      <div style="line-height: 0; font-size: 0; mso-line-height-rule: exactly;">
      <table cellpadding="0" cellspacing="0" border="0" width="100%" height="4" style="width: 100%; height: 4px; border-spacing: 0; border-collapse: collapse; line-height: 0; font-size: 0; mso-line-height-rule: exactly;">
        <tr height="4" style="height: 4px; line-height: 0; mso-line-height-rule: exactly;">
          <td width="20%" height="4" bgcolor="#78a22f" valign="top" style="width: 20%; height: 4px; max-height: 4px; min-height: 4px; background-color: #78a22f; font-size: 0; line-height: 0; mso-line-height-rule: exactly; padding: 0; border: 0; overflow: hidden;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="" width="1" height="4" style="display: block; width: 1px; height: 4px; max-height: 4px; border: 0; opacity: 0;" /></td>
          <td width="20%" height="4" bgcolor="#e74c30" valign="top" style="width: 20%; height: 4px; max-height: 4px; min-height: 4px; background-color: #e74c30; font-size: 0; line-height: 0; mso-line-height-rule: exactly; padding: 0; border: 0; overflow: hidden;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="" width="1" height="4" style="display: block; width: 1px; height: 4px; max-height: 4px; border: 0; opacity: 0;" /></td>
          <td width="20%" height="4" bgcolor="#8b6baf" valign="top" style="width: 20%; height: 4px; max-height: 4px; min-height: 4px; background-color: #8b6baf; font-size: 0; line-height: 0; mso-line-height-rule: exactly; padding: 0; border: 0; overflow: hidden;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="" width="1" height="4" style="display: block; width: 1px; height: 4px; max-height: 4px; border: 0; opacity: 0;" /></td>
          <td width="20%" height="4" bgcolor="#c5c946" valign="top" style="width: 20%; height: 4px; max-height: 4px; min-height: 4px; background-color: #c5c946; font-size: 0; line-height: 0; mso-line-height-rule: exactly; padding: 0; border: 0; overflow: hidden;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="" width="1" height="4" style="display: block; width: 1px; height: 4px; max-height: 4px; border: 0; opacity: 0;" /></td>
          <td width="20%" height="4" bgcolor="#3fa7c9" valign="top" style="width: 20%; height: 4px; max-height: 4px; min-height: 4px; background-color: #3fa7c9; font-size: 0; line-height: 0; mso-line-height-rule: exactly; padding: 0; border: 0; overflow: hidden;"><img src="https://allenbenedikt.github.io/spsi-email-assets/spsi-logo.png?raw=true" alt="" width="1" height="4" style="display: block; width: 1px; height: 4px; max-height: 4px; border: 0; opacity: 0;" /></td>
        </tr>
      </table>
      </div>
      <p style="margin: 10px 0 0 0;"><strong>Confidentiality Notice:</strong> All information pertaining to this email contains confidential information intended only for the use of the recipient(s) named in the header text. If you are not the intended recipient, you are hereby notified that any disclosure, copying, distribution, or the taking of any action in reliance on the contents of this emailed information except its direct delivery to the person named above is strictly prohibited. If you have received this email in error, please notify us immediately by replying to this email and delete all copies of this message. This message is protected by applicable legal privileges and is confidential.</p>
    </td>
  </tr>
</table>
"""


def has_value(val):
    """True if a spreadsheet cell has a real value (not NaN, not empty string)."""
    return pd.notna(val) and str(val).strip() != ""


def tel_digits(phone):
    """Strip all non-digit characters for use inside tel: links."""
    if not has_value(phone):
        return ""
    return re.sub(r"\D", "", str(phone))


def format_title(title):
    if has_value(title):
        return f'<p style="margin: 0; font-size: 14px;">{title}</p>'
    return ""


def format_extension(extension):
    if has_value(extension):
        return f"<strong>Ext.</strong> {int(float(extension))}"
    return ""


def format_mobile(mobile):
    """Optional Mobile row, placed ABOVE the Main row when present."""
    if has_value(mobile):
        return (
            f'<p style="margin: 6px 0 0; font-size: 13px; color: #000000;">'
            f'<strong>Mobile:</strong> '
            f'<a href="tel:{tel_digits(mobile)}" style="color: #000000; text-decoration: none;">{mobile}</a>'
            f'</p>'
        )
    return ""


def format_direct(phone_direct):
    """Optional Direct phone, appended inline after Main with a leading separator."""
    if has_value(phone_direct):
        return (
            f'| <strong>Direct:</strong> '
            f'<a href="tel:{tel_digits(phone_direct)}" style="color: #000000; text-decoration: none;">{phone_direct}</a>'
        )
    return ""


def format_email(email):
    if has_value(email):
        return (
            f'<p style="margin: 0; font-size: 13px; color: #000000;">'
            f'<strong>Email:</strong> '
            f'<a href="mailto:{email}" style="color: #000000; text-decoration: none;">{email}</a>'
            f'</p>'
        )
    return ""


# Output directory
output_folder = "2026 Signatures"
os.makedirs(output_folder, exist_ok=True)

# Generate HTML signatures
for _, row in df_uploaded.iterrows():
    title_section = format_title(row.get("Title"))
    mobile_section = format_mobile(row.get("PhoneMobile"))
    extension_info = format_extension(row.get("Extension"))
    direct_section = format_direct(row.get("PhoneDirect"))
    email_section = format_email(row.get("Email"))

    phone_main = row.get("PhoneMain", "")
    phone_main_tel = tel_digits(phone_main)

    html_content = html_template.format(
        name=row["Name"],
        title_section=title_section,
        mobile_section=mobile_section,
        phone_main=phone_main,
        phone_main_tel=phone_main_tel,
        extension_info=extension_info,
        direct_section=direct_section,
        email_section=email_section,
        address1=row.get("Address1", ""),
        address2=row.get("Address2", ""),
    )

    file_name = f"{row['Name'].replace(' ', '_')}_signature.html"
    with open(os.path.join(output_folder, file_name), "w", encoding="utf-8") as f:
        f.write(html_content)

print(f"Signatures generated in {output_folder}")
