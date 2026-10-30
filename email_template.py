# -*- coding: utf-8 -*-
"""
Email Template - Mail merge engine.
Supports {field_name} placeholders in email body template.
"""


def format_number(value):
    """Format number with comma separators."""
    if value is None or value == '' or value == 0:
        return ''
    try:
        num = float(str(value).replace(',', ''))
        if num == int(num):
            return f"{int(num):,}"
        return f"{num:,.0f}"
    except (ValueError, TypeError):
        return str(value)


def generate_debt_table_html(debt_items, company_name=''):
    """Generate HTML debt table from data items."""
    total_phai_thu = sum(
        float(str(item.get('so_tien_phai_thu', 0) or 0).replace(',', ''))
        for item in debt_items
    )
    total_da_thu = sum(
        float(str(item.get('da_thu_tien', 0) or 0).replace(',', ''))
        for item in debt_items
    )
    total_con_lai = sum(
        float(str(item.get('con_lai_phai_thu', 0) or 0).replace(',', ''))
        for item in debt_items
    )

    hdr = "background:ivory;border:solid windowtext 1.0pt;font-family:'Times New Roman',serif;font-weight:bold;color:black;text-align:center;padding:0in 5.4pt 0in 5.4pt;"
    cs = "font-family:'Times New Roman',serif;color:black;border:solid windowtext 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt;"
    cr = cs + "text-align:right;"
    cc = cs + "text-align:center;"

    html = f'''<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=1015 style='width:760.95pt;border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0in 0in 0in 0in'>
<tr style='height:11.2pt'>
  <td rowspan=2 style="{hdr}width:216.15pt;"><b>NỘI DUNG</b></td>
  <td colspan=2 style="{hdr}width:98.75pt;border-left:none;"><b>HÓA ĐƠN</b></td>
  <td rowspan=2 style="{hdr}width:77.75pt;border-left:none;"><b>SỐ TIỀN<br>PHẢI THU</b></td>
  <td rowspan=2 style="{hdr}width:67.75pt;border-left:none;"><b>THỜI<br>HẠN TT<br>(Ngày)</b></td>
  <td rowspan=2 style="{hdr}width:59.35pt;border-left:none;"><b>HẠN<br>THU</b></td>
  <td rowspan=2 style="{hdr}width:77.75pt;border-left:none;"><b>ĐÃ<br>THU TIỀN</b></td>
  <td rowspan=2 style="{hdr}width:77.75pt;border-left:none;"><b style="color:blue;">CÒN LẠI<br>PHẢI THU</b></td>
  <td rowspan=2 style="{hdr}width:85.7pt;border-left:none;"><b>QUÁ HẠN<br>TT (Ngày)</b></td>
</tr>
<tr style='height:11.2pt'>
  <td style="{hdr}width:39.35pt;border-top:none;border-left:none;"><b>SỐ</b></td>
  <td style="{hdr}width:59.4pt;border-top:none;border-left:none;"><b>NGÀY</b></td>
</tr>
<tr style='height:11.2pt'>
  <td valign=top style="{cs}border-top:none;font-weight:bold;"><b>{company_name}</b></td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">&nbsp;</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">&nbsp;</td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;font-weight:bold;"><b>{format_number(total_phai_thu)}</b></td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">&nbsp;</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">&nbsp;</td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;font-weight:bold;"><b>{format_number(total_da_thu)}</b></td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;color:blue;font-weight:bold;"><b>{format_number(total_con_lai)}</b></td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">&nbsp;</td>
</tr>'''

    for item in debt_items:
        html += f'''
<tr style='height:11.2pt'>
  <td valign=top style="{cs}border-top:none;">{item.get('noi_dung', '')}</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">{item.get('hd_so', '')}</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">{item.get('hd_ngay', '')}</td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;">{format_number(item.get('so_tien_phai_thu', ''))}</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">{item.get('thoi_han_tt', '')}</td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">{item.get('han_thu', '')}</td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;">{format_number(item.get('da_thu_tien', ''))}</td>
  <td nowrap valign=top style="{cr}border-top:none;border-left:none;color:blue;font-weight:bold;"><b>{format_number(item.get('con_lai_phai_thu', ''))}</b></td>
  <td nowrap valign=top style="{cc}border-top:none;border-left:none;">{item.get('qua_han_tt', '')}</td>
</tr>'''

    html += '\n</table>'
    return html


def merge_template(template_text, fields, debt_items=None, company_name=''):
    """
    Mail merge: replace {field_name} placeholders in template with actual values.
    
    Special field: {debt_table} is auto-generated from debt_items.
    
    Args:
        template_text: email body template with {placeholders}
        fields: dict of field_name -> value
        debt_items: list of debt data dicts (for {debt_table})
        company_name: company name for debt table header row
    
    Returns:
        str: merged HTML email body
    """
    result = template_text

    # Generate debt table if needed
    if '{debt_table}' in result and debt_items is not None:
        table_html = generate_debt_table_html(debt_items, company_name)
        result = result.replace('{debt_table}', table_html)

    # Replace all other fields
    for key, val in fields.items():
        placeholder = '{' + key + '}'
        result = result.replace(placeholder, str(val) if val else '')

    # Wrap in HTML - set margin:0 on p tags to prevent Outlook double-spacing
    html = f'''<html>
<head><meta charset="utf-8">
<style>p {{ margin:0; padding:0; line-height:1.5; }}</style>
</head>
<body style="font-family:'Times New Roman',serif;font-size:11pt;color:black;line-height:1.5;">
{result}
</body>
</html>'''
    return html


# Default email body template (compact spacing for Outlook)
DEFAULT_TEMPLATE = """<p><b>Kính gửi: {customer_name}</b></p>
<p>&nbsp;</p>
<p>{company_name} chân thành cảm ơn Quý khách hàng đã đặt hàng các sản phẩm của chúng tôi trong thời gian qua.</p>
<p>&nbsp;</p>
<p>Nay chúng tôi gửi đến Quý khách hàng tổng công nợ quá hạn tính đến thời điểm <b>{debt_date}</b> như sau:</p>

{debt_table}

<p>Rất mong nhận được sự hợp tác từ Quý khách hàng trong việc <b style="color:red;"><i><u>thanh toán nợ quá hạn trên trước ngày {payment_deadline}.</u></i></b> Trong trường hợp Quý khách đã thực hiện thanh toán qua ngân hàng, Quý khách vui lòng gửi cho chúng tôi ảnh chụp/file UNC để chúng tôi thực hiện tra soát với ngân hàng và ghi giảm công nợ.</p>
<p>&nbsp;</p>
<p>Chúng tôi luôn mong muốn duy trì quan hệ hợp tác lâu dài và tin rằng việc thanh toán các khoản nợ là cơ sở để hai bên tiếp tục phát triển bền vững. Mong sớm nhận được phản hồi từ Quý khách để tránh ảnh hưởng đến hoạt động kinh doanh của cả hai bên.</p>
<p>&nbsp;</p>
<p>Xin trân trọng cảm ơn!</p>

{signature}"""

