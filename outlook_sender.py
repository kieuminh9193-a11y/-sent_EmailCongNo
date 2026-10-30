# -*- coding: utf-8 -*-
"""
Outlook Sender - Send email via Outlook COM automation.
Requires: Outlook desktop installed, pywin32.
"""
import os


def _init_com():
    """Initialize COM in current thread (required for Flask threads)."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def create_outlook_email(to, cc, subject, html_body, attachments=None,
                          send_mode='display', from_account=None, bcc=None):
    """
    Create and send email via Outlook COM.

    Args:
        to: recipient email(s), separated by ;
        cc: CC email(s), separated by ;
        subject: email subject
        html_body: HTML email body
        attachments: list of file paths to attach
        send_mode: 'display' (show in Outlook) or 'send' (send immediately)
        bcc: BCC email(s), separated by ;
        from_account: sender email or alias address

    Returns:
        dict: {'success': bool, 'message': str}
    """
    try:
        _init_com()
        import win32com.client

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem

        # Set From account if specified
        if from_account:
            try:
                namespace = outlook.GetNamespace("MAPI")
                matched = False
                # First: try to match directly to an account
                for account in namespace.Accounts:
                    if account.SmtpAddress.lower() == from_account.lower():
                        mail._oleobj_.Invoke(
                            *(64209, 0, 8, 0, account)
                        )
                        matched = True
                        break

                # If not matched directly, it might be an alias or shared mailbox
                if not matched:
                    # Find the parent account that owns this alias
                    parent_account = _find_parent_account_for_alias(
                        namespace, from_account
                    )
                    if parent_account:
                        # Set the parent account as SendUsingAccount
                        mail._oleobj_.Invoke(
                            *(64209, 0, 8, 0, parent_account)
                        )
                    # Set SentOnBehalfOfName to the alias
                    mail.SentOnBehalfOfName = from_account
            except Exception:
                pass

        mail.To = to or ''
        mail.CC = cc or ''
        mail.BCC = bcc or ''
        mail.Subject = subject or ''
        mail.HTMLBody = html_body

        # Attach files
        if attachments:
            for filepath in attachments:
                if filepath and os.path.exists(filepath):
                    mail.Attachments.Add(os.path.abspath(filepath))

        if send_mode == 'send':
            mail.Send()
            return {'success': True, 'message': f'Email sent successfully to {to}'}
        else:
            mail.Display(False)
            return {'success': True, 'message': f'Email opened in Outlook for {to}'}

    except ImportError:
        return {'success': False, 'message': 'Error: pywin32 not installed. Run: pip install pywin32'}
    except Exception as e:
        return {'success': False, 'message': f'Error creating email: {str(e)}'}


def _find_parent_account_for_alias(namespace, alias_email):
    """Find the Outlook Account object that owns the given alias email."""
    try:
        for account in namespace.Accounts:
            aliases = _get_account_aliases(namespace, account)
            for a in aliases:
                if a.lower() == alias_email.lower():
                    return account
    except Exception:
        pass
    return None


def _get_account_aliases(namespace, account):
    """
    Get alias (proxy) email addresses for an Outlook account.
    Works for Exchange / Office 365 accounts that have proxy addresses.
    Returns a list of alias email strings (excluding the primary SmtpAddress).
    """
    aliases = []
    PR_EMS_AB_PROXY_ADDRESSES = "http://schemas.microsoft.com/mapi/proptag/0x800F101F"

    # Method 1: Try using CurrentUser.AddressEntry (highly reliable for active accounts)
    try:
        ae = account.CurrentUser.AddressEntry
        if ae:
            prop = ae.PropertyAccessor.GetProperty(PR_EMS_AB_PROXY_ADDRESSES)
            if prop:
                for val in prop:
                    val_str = str(val)
                    if val_str.lower().startswith("smtp:"):
                        email = val_str[5:].strip()
                        if email.lower() != account.SmtpAddress.lower() and email.lower() not in [a.lower() for a in aliases]:
                            aliases.append(email)
                if aliases:
                    return aliases
    except Exception:
        pass

    # Method 2: Fallback to CreateRecipient (for secondary or other types of accounts)
    try:
        recipient = namespace.CreateRecipient(account.SmtpAddress)
        recipient.Resolve()
        if recipient.Resolved:
            ae = recipient.AddressEntry
            if ae:
                prop = ae.PropertyAccessor.GetProperty(PR_EMS_AB_PROXY_ADDRESSES)
                if prop:
                    for val in prop:
                        val_str = str(val)
                        if val_str.lower().startswith("smtp:"):
                            email = val_str[5:].strip()
                            if email.lower() != account.SmtpAddress.lower() and email.lower() not in [a.lower() for a in aliases]:
                                aliases.append(email)
    except Exception:
        pass

    return aliases


def check_outlook_available():
    """Check if Outlook is available and return account list with aliases."""
    try:
        _init_com()
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts = []
        for account in namespace.Accounts:
            acc_info = {
                'email': account.SmtpAddress,
                'name': account.DisplayName,
                'aliases': _get_account_aliases(namespace, account)
            }
            accounts.append(acc_info)
        return {'available': True, 'accounts': accounts}
    except Exception as e:
        return {'available': False, 'error': str(e), 'accounts': []}

