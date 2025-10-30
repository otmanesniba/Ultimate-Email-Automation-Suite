import asyncio
import smtplib
import imaplib
import email
from pyppeteer import launch
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
import re
import os
import time
import pyfiglet
from termcolor import colored
from getpass import getpass
import sys
from datetime import datetime, timedelta

class UltimateEmailAutomationSuite:
    def __init__(self):
        self.collected_emails = []
        self.current_step = 0
        self.total_steps = 4
        self.session_stats = {
            'emails_collected': 0,
            'emails_sent': 0,
            'successful_sends': 0,
            'failed_sends': 0,
            'start_time': None
        }
        self.imap_connection = None
        self.found_emails_data = []
        # Set the custom save path
        self.custom_save_path = r"C:\Users\otmane sniba\Desktop\Stage script"
        
    def clear_screen(self):
        """Clear the terminal screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def print_header(self, title):
        """Print a beautiful header"""
        print(colored("╔" + "═" * 78 + "╗", "cyan"))
        title_line = f"║ {title:^76} ║"
        print(colored(title_line, "cyan"))
        print(colored("╚" + "═" * 78 + "╝", "cyan"))
    
    def print_section(self, title):
        """Print a section header"""
        print(colored(f"\n┌─ {title} ", "yellow") + colored("─" * (76 - len(title)) + "┐", "yellow"))
    
    def print_section_end(self):
        """Print section end"""
        print(colored("└" + "─" * 78 + "┘", "yellow"))
    
    def print_status(self, message, status_type="info"):
        """Print status messages with icons"""
        icons = {
            "success": "✅",
            "error": "❌",
            "warning": "⚠️",
            "info": "🔵",
            "progress": "🔄",
            "email": "📧",
            "file": "📁",
            "link": "🔗",
            "delete": "🗑️",
            "search": "🔍",
            "import": "📥",
            "export": "📤"
        }
        
        colors = {
            "success": "green",
            "error": "red",
            "warning": "yellow",
            "info": "blue",
            "progress": "cyan",
            "email": "magenta",
            "file": "cyan",
            "link": "blue",
            "delete": "red",
            "search": "yellow",
            "import": "green",
            "export": "blue"
        }
        
        icon = icons.get(status_type, "➡️")
        color = colors.get(status_type, "white")
        print(colored(f"  {icon} {message}", color))
    
    def print_progress_bar(self, iteration, total, prefix='', suffix='', length=50, fill='█'):
        """Create a progress bar"""
        percent = ("{0:.1f}").format(100 * (iteration / float(total)))
        filled_length = int(length * iteration // total)
        bar = fill * filled_length + '░' * (length - filled_length)
        print(f'\r  {prefix} |{bar}| {percent}% {suffix}', end='\r')
        if iteration == total:
            print()
    
    def print_banner(self):
        """Display the enhanced main banner"""
        self.clear_screen()
        
        # Main banner with gradient effect
        ascii_banner = pyfiglet.figlet_format("ULTIMATE EMAIL SUITE", font="small")
        lines = ascii_banner.split('\n')
        
        colors = ['cyan', 'blue', 'magenta', 'blue', 'cyan']
        for i, line in enumerate(lines):
            color = colors[i % len(colors)]
            print(colored(line.center(80), color))
        
        # Subtitle
        print(colored("🚀 PROFESSIONAL EMAIL AUTOMATION & MANAGEMENT PLATFORM".center(80), "yellow"))
        print(colored("Made by Otmane Sniba".center(80), "cyan"))
        print()
        
        # Session info
        if self.session_stats['start_time']:
            duration = datetime.now() - self.session_stats['start_time']
            print(colored(f"🕒 Session active for: {duration}", "blue"))
        
        print(colored("═" * 80, "blue"))
    
    def display_dashboard(self):
        """Display the main dashboard with statistics"""
        self.print_banner()
        
        # Statistics panel
        self.print_header("DASHBOARD OVERVIEW")
        
        print(colored("\n📊 CURRENT STATISTICS:", "magenta"))
        print(colored("┌───────────────────┬──────────┬───────────────────┐", "cyan"))
        print(colored("│      Metric       │  Count   │      Status       │", "cyan"))
        print(colored("├───────────────────┼──────────┼───────────────────┤", "cyan"))
        
        stats = [
            ("Emails Collected", len(self.collected_emails), 
             "✅ Ready" if self.collected_emails else "⏳ Waiting"),
            ("Successful Sends", self.session_stats['successful_sends'],
             "🎉 Excellent" if self.session_stats['successful_sends'] > 0 else "➡️ Pending"),
            ("Failed Sends", self.session_stats['failed_sends'],
             "❌ Needs Review" if self.session_stats['failed_sends'] > 0 else "✅ Perfect")
        ]
        
        for metric, count, status in stats:
            metric_col = colored(f"│ {metric:17} ", "white")
            count_col = colored(f"│ {count:8} ", "green")
            status_col = colored(f"│ {status:17} │", "yellow")
            print(metric_col + count_col + status_col)
        
        print(colored("└───────────────────┴──────────┴───────────────────┘", "cyan"))
        
        # Display available actions with fixed table formatting
        print(colored("\n🎯 AVAILABLE ACTIONS:", "magenta"))
        print(colored("┌─────┬──────────────────┬─────────────────────────────────────────────┐", "cyan"))
        print(colored("│ Key │      Action      │                Description                  │", "cyan"))
        print(colored("├─────┼──────────────────┼─────────────────────────────────────────────┤", "cyan"))
        
        menu_options = [
            ("1", "📧 Collect Emails", "Extract emails from websites"),
            ("2", "🚀 Send Campaign", "Send emails to collected addresses"),
            ("3", "📥 Import Emails", "Load emails from existing file"),
            ("4", "👁️ View Emails", "Browse collected email addresses"),
            ("5", "🔄 Clear Data", "Reset collected emails"),
            ("6", "🔍 Email Management", "Search & manage sent emails"),
            ("7", "📊 Session Report", "View detailed statistics"),
            ("8", "❌ Exit", "Close the application")
        ]
        
        for key, action, description in menu_options:
            key_col = colored(f"│  {key}  ", "yellow")
            action_col = colored(f"│ {action:16} ", "green")
            desc_col = colored(f"│ {description:43} │", "white")
            print(key_col + action_col + desc_col)
        
        print(colored("└─────┴──────────────────┴─────────────────────────────────────────────┘", "cyan"))
    
    def display_menu(self):
        """Display the enhanced main interactive menu"""
        if self.session_stats['start_time'] is None:
            self.session_stats['start_time'] = datetime.now()
            
        while True:
            self.display_dashboard()
            
            choice = input(colored("\n🎮 Select an option (1-8): ", "cyan")).strip()
            
            if choice == "1":
                self.collect_emails_flow()
            elif choice == "2":
                self.send_emails_flow()
            elif choice == "3":
                self.import_emails_flow()
            elif choice == "4":
                self.view_collected_emails()
            elif choice == "5":
                self.clear_emails()
            elif choice == "6":
                self.email_management_flow()
            elif choice == "7":
                self.show_session_report()
            elif choice == "8":
                self.exit_gracefully()
            else:
                self.print_status("Invalid option! Please try again.", "error")
                self.press_enter_to_continue()

    def read_urls_from_file(self, file_path):
        """Read URLs from a file with enhanced error handling"""
        try:
            if os.path.exists(file_path):
                with open(file_path, "r") as file:
                    urls = [line.strip() for line in file.readlines() if line.strip()]
                return urls
            else:
                self.print_status(f"File not found: {file_path}", "error")
                return []
        except Exception as e:
            self.print_status(f"Error reading file: {e}", "error")
            return []

    def import_emails_flow(self):
        """Import emails from existing files and optionally send immediately"""
        self.print_banner()
        self.print_header("EMAIL IMPORT WIZARD")
        
        self.print_section("FILE SELECTION")
        email_file_path = input(colored("📁 Drag and drop your email list file here: ", "yellow")).strip().strip('"')
        
        if not email_file_path or not os.path.exists(email_file_path):
            self.print_status("File not found or invalid path.", "error")
            self.press_enter_to_continue()
            return
        
        imported_emails = self.read_emails_from_file(email_file_path)
        
        if not imported_emails:
            self.print_status("No valid emails found in the file.", "warning")
            self.press_enter_to_continue()
            return
        
        self.print_status(f"Successfully imported {len(imported_emails)} emails from file", "success")
        
        # Show preview
        self.print_section("IMPORTED EMAILS PREVIEW")
        for i, email in enumerate(imported_emails[:8], 1):
            print(colored(f"  {i:2d}. {email}", "cyan"))
        if len(imported_emails) > 8:
            print(colored(f"  ... and {len(imported_emails) - 8} more", "blue"))
        
        # Ask user what to do
        self.print_section("ACTION SELECTION")
        print(colored("🎯 What would you like to do with these emails?", "magenta"))
        print(colored("1. Add to current collection", "cyan"))
        print(colored("2. Replace current collection", "yellow"))
        print(colored("3. Send campaign immediately", "green"))
        print(colored("4. Just view (no action)", "blue"))
        
        action = input(colored("\nSelect action (1-4): ", "cyan")).strip()
        
        if action == "1":
            # Add to current collection
            before_count = len(self.collected_emails)
            self.collected_emails.extend(imported_emails)
            self.collected_emails = list(set(self.collected_emails))  # Remove duplicates
            added_count = len(self.collected_emails) - before_count
            self.print_status(f"Added {added_count} new emails. Total: {len(self.collected_emails)}", "success")
            self.press_enter_to_continue()
            
        elif action == "2":
            # Replace current collection
            self.collected_emails = imported_emails
            self.print_status(f"Replaced collection with {len(imported_emails)} emails", "success")
            self.press_enter_to_continue()
            
        elif action == "3":
            # Send campaign immediately
            self.collected_emails = imported_emails
            self.print_status(f"Loaded {len(imported_emails)} emails for immediate sending", "success")
            self.press_enter_to_continue()
            # Launch send flow directly
            self.send_emails_flow()
            
        elif action == "4":
            # Just view
            self.print_status("No changes made to collection.", "info")
            self.press_enter_to_continue()
        else:
            self.print_status("Invalid action selected.", "warning")
            self.press_enter_to_continue()

    def read_emails_from_file(self, file_path):
        """Read emails from various file formats"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read()
            
            # Extract emails using regex pattern
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            emails = re.findall(email_pattern, content)
            
            # Remove duplicates and return
            unique_emails = list(set(emails))
            return unique_emails
            
        except Exception as e:
            self.print_status(f"Error reading file: {e}", "error")
            return []

    def quick_send_flow(self, emails):
        """Quick send flow for imported emails"""
        self.print_banner()
        self.print_header("QUICK EMAIL CAMPAIGN")
        
        self.print_section("CAMPAIGN SETUP")
        self.print_status(f"Ready to send to {len(emails)} imported emails", "success")
        
        # Get sender credentials
        sender_email = input(colored("📧 Enter your email address: ", "yellow")).strip()
        password = getpass(colored("🔒 Enter your email password: ", "yellow")).strip()
        
        # Get email content
        self.print_section("EMAIL CONTENT")
        letter_path = input(colored("📝 Drag and drop your Letter.txt file here: ", "yellow")).strip().strip('"')
        
        try:
            with open(letter_path, "r", encoding="utf-8") as letter_file:
                subject = letter_file.readline().strip()
                body = letter_file.read().strip()
            self.print_status("Email content loaded successfully!", "success")
        except FileNotFoundError:
            self.print_status("Error: Letter file not found.", "error")
            return
        
        # Quick attachment setup
        self.print_section("ATTACHMENTS (Optional)")
        print(colored("💡 You can skip attachments by pressing Enter", "blue"))
        cv_path = input(colored("📎 Drag and drop your CV file (or press Enter to skip): ", "yellow")).strip().strip('"')
        demand_path = input(colored("📎 Drag and drop your Demand file (or press Enter to skip): ", "yellow")).strip().strip('"')
        
        files_to_attach = []
        for path, name in [(cv_path, "CV"), (demand_path, "Demand Letter")]:
            if path and os.path.exists(path):
                files_to_attach.append(path)
                self.print_status(f"{name} attached: {os.path.basename(path)}", "success")
        
        # Quick confirmation
        self.print_section("QUICK CONFIRMATION")
        print(colored(f"🎯 Ready to send to {len(emails)} recipients", "magenta"))
        print(colored(f"📝 Subject: {subject}", "cyan"))
        print(colored(f"📎 Attachments: {len(files_to_attach)}", "cyan"))
        
        confirm = input(colored("\n🚀 Start sending now? (y/n): ", "red")).lower()
        if confirm == 'y':
            self.send_emails(sender_email, password, emails, subject, body, files_to_attach)
        else:
            self.print_status("Quick send cancelled.", "warning")
        
        self.press_enter_to_continue()

    def email_management_flow(self):
        """Enhanced email management functionality"""
        self.print_banner()
        self.print_header("EMAIL MANAGEMENT CENTER")
        
        self.print_section("GMAIL ACCOUNT ACCESS")
        print(colored("🔐 This feature requires access to your Gmail account", "yellow"))
        print(colored("   to search and manage sent emails.", "yellow"))
        
        email_user = input(colored("\n📧 Enter your Gmail address: ", "yellow")).strip()
        email_pass = getpass(colored("🔒 Enter your Gmail password (App Password if 2FA): ", "yellow")).strip()
        
        try:
            # Connect to the IMAP server
            self.imap_connection = imaplib.IMAP4_SSL("imap.gmail.com")
            self.imap_connection.login(email_user, email_pass)
            self.print_status("Gmail login successful!", "success")
            
            while True:
                self.print_banner()
                self.print_header("EMAIL MANAGEMENT CENTER")
                
                print(colored("\n🔧 MANAGEMENT OPTIONS:", "magenta"))
                print(colored("1. 🔍 Search Sent Emails by Subject", "cyan"))
                print(colored("2. 💾 Save Found Emails to File", "cyan"))
                print(colored("3. 🗑️  Delete Emails from Sent Folder", "cyan"))
                print(colored("4. 📥 Import & Send Found Emails", "green"))
                print(colored("5. ↩️  Return to Main Menu", "yellow"))
                
                choice = input(colored("\n🎮 Select an option (1-5): ", "cyan")).strip()
                
                if choice == "1":
                    found_emails = self.search_sent_emails_by_subject(email_user)
                    if found_emails:
                        # Offer to send immediately
                        send_now = input(colored("\n🚀 Send to these emails now? (y/n): ", "green")).lower()
                        if send_now == 'y':
                            self.quick_send_flow(found_emails)
                elif choice == "2":
                    self.save_found_emails_to_file()
                elif choice == "3":
                    self.delete_sent_emails()
                elif choice == "4":
                    self.import_and_send_found_emails()
                elif choice == "5":
                    break
                else:
                    self.print_status("Invalid option! Please try again.", "error")
                
                self.press_enter_to_continue()
                
        except Exception as e:
            self.print_status(f"Failed to login: {e}", "error")
            self.press_enter_to_continue()
        finally:
            if self.imap_connection:
                try:
                    self.imap_connection.logout()
                    self.print_status("Gmail connection closed.", "info")
                except:
                    pass

    def import_and_send_found_emails(self):
        """Import emails from file and send immediately"""
        self.print_section("QUICK IMPORT & SEND")
        
        email_file_path = input(colored("📁 Drag and drop your email list file here: ", "yellow")).strip().strip('"')
        
        if not email_file_path or not os.path.exists(email_file_path):
            self.print_status("File not found or invalid path.", "error")
            return
        
        imported_emails = self.read_emails_from_file(email_file_path)
        
        if not imported_emails:
            self.print_status("No valid emails found in the file.", "warning")
            return
        
        self.print_status(f"Found {len(imported_emails)} emails in file", "success")
        self.quick_send_flow(imported_emails)

    def search_sent_emails_by_subject(self, email_user):
        """Search sent emails by subject and extract recipient emails"""
        self.print_section("SEARCH SENT EMAILS")
        
        try:
            # Select the 'Sent Mail' folder
            self.imap_connection.select('"[Gmail]/Sent Mail"')
            
            subject_start = input(colored("🔍 Enter the first words of the subject to search: ", "yellow")).strip()
            
            if not subject_start:
                self.print_status("No search terms provided.", "warning")
                return []
            
            # Search for emails with subjects starting with the given words
            search_query = f'SUBJECT "{subject_start}"'
            status, messages = self.imap_connection.search(None, search_query)
            
            if status != "OK" or not messages[0]:
                self.print_status(f"No emails found with subject starting with '{subject_start}'.", "warning")
                return []
            
            email_ids = messages[0].split()
            self.print_status(f"Found {len(email_ids)} email(s) with subject starting with '{subject_start}'.", "success")
            
            found_emails = []
            
            for i, email_id in enumerate(email_ids):
                status, msg_data = self.imap_connection.fetch(email_id, "(RFC822)")
                if status != "OK":
                    continue
                
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        
                        # Get the recipient(s) from the "To" field
                        to_field = msg.get("To", "")
                        
                        if to_field:
                            recipients = [addr.strip() for addr in to_field.split(",")]
                            
                            for recipient in recipients:
                                # Extract email from recipient string
                                email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', recipient)
                                if email_match:
                                    found_email = email_match.group()
                                    if found_email != email_user and found_email not in found_emails:
                                        found_emails.append(found_email)
                                        self.print_status(f"Found: {found_email}", "email")
            
            if found_emails:
                self.found_emails_data = found_emails
                self.print_status(f"Extracted {len(found_emails)} unique recipient emails.", "success")
                return found_emails
            else:
                self.print_status("No recipient emails found in the search results.", "warning")
                return []
                
        except Exception as e:
            self.print_status(f"Error searching emails: {e}", "error")
            return []

    def save_emails_to_file_direct(self, emails, filename_prefix):
        """Save emails directly to file using custom path"""
        try:
            # Use custom save path instead of desktop
            if not os.path.exists(self.custom_save_path):
                os.makedirs(self.custom_save_path)
                
            filename = f"{filename_prefix}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
            file_path = os.path.join(self.custom_save_path, filename)
            
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(f"# Emails Found - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# Total: {len(emails)}\n")
                f.write("#" * 50 + "\n\n")
                for email_addr in emails:
                    f.write(f"{email_addr}\n")
            
            self.print_status(f"Emails saved to: {file_path}", "success")
            return file_path
        except Exception as e:
            self.print_status(f"Error saving file: {e}", "error")
            return None

    def delete_emails_from_sent(self, email_ids):
        """Delete emails from sent folder"""
        self.print_section("DELETE SENT EMAILS")
        
        if not email_ids:
            self.print_status("No email IDs provided for deletion.", "warning")
            return
        
        print(colored("⚠️  WARNING: This will permanently delete emails from your Sent folder!", "red"))
        confirm = input(colored("🔴 Type 'DELETE' to confirm: ", "red")).strip()
        
        if confirm.upper() != "DELETE":
            self.print_status("Deletion cancelled.", "info")
            return
        
        try:
            for email_id in email_ids:
                self.imap_connection.store(email_id, "+FLAGS", "\\Deleted")
            self.imap_connection.expunge()
            self.print_status(f"Successfully deleted {len(email_ids)} email(s) from Sent folder.", "success")
        except Exception as e:
            self.print_status(f"Error deleting emails: {e}", "error")

    def save_found_emails_to_file(self):
        """Save currently found emails to file"""
        if not hasattr(self, 'found_emails_data') or not self.found_emails_data:
            self.print_status("No found emails to save. Please search for emails first.", "warning")
            return
        
        filename = input(colored("📝 Enter filename prefix (or press Enter for default): ", "yellow")).strip()
        if not filename:
            filename = "managed_emails"
        
        self.save_emails_to_file_direct(self.found_emails_data, filename)

    def delete_sent_emails(self):
        """Direct interface for deleting sent emails"""
        self.print_section("DELETE SENT EMAILS DIRECTLY")
        
        try:
            self.imap_connection.select('"[Gmail]/Sent Mail"')
            
            subject_start = input(colored("🔍 Enter subject words to find emails to delete: ", "yellow")).strip()
            
            if not subject_start:
                self.print_status("No search terms provided.", "warning")
                return
            
            search_query = f'SUBJECT "{subject_start}"'
            status, messages = self.imap_connection.search(None, search_query)
            
            if status != "OK" or not messages[0]:
                self.print_status(f"No emails found with subject containing '{subject_start}'.", "warning")
                return
            
            email_ids = messages[0].split()
            self.print_status(f"Found {len(email_ids)} email(s) to delete.", "warning")
            
            self.delete_emails_from_sent(email_ids)
            
        except Exception as e:
            self.print_status(f"Error in deletion process: {e}", "error")

    def collect_emails_flow(self):
        """Enhanced email collection process"""
        self.print_banner()
        self.print_header("EMAIL COLLECTION WIZARD")
        
        self.print_section("INPUT SOURCE")
        file_path = input(colored("🎯 Please enter the path to your URLs file: ", "magenta")).strip()
        
        urls = self.read_urls_from_file(file_path)
        if not urls:
            self.print_status("No URLs found. Please check your file.", "error")
            self.press_enter_to_continue()
            return
        
        self.print_status(f"Found {len(urls)} URLs to process", "success")
        
        # Scraping progress
        self.print_section("SCRAPING PROGRESS")
        emails = asyncio.run(self.scrape_emails_with_pyppeteer(urls))
        
        if emails:
            self.collected_emails = list(set(emails))
            self.session_stats['emails_collected'] += len(self.collected_emails)
            
            self.print_section("COLLECTION RESULTS")
            self.print_status(f"Successfully collected {len(self.collected_emails)} unique emails!", "success")
            
            # Save to file automatically using custom path
            saved_path = self.save_emails_to_file()
            self.print_status(f"Emails saved to: {saved_path}", "file")
            
            # Ask if user wants to send immediately
            send_now = input(colored("\n🚀 Send campaign to these emails now? (y/n): ", "green")).lower()
            if send_now == 'y':
                self.send_emails_flow()
            else:
                self.press_enter_to_continue()
        else:
            self.print_status("No emails found during scraping process.", "warning")
            self.press_enter_to_continue()

    async def scrape_emails_with_pyppeteer(self, urls):
        """Enhanced email scraping with better progress tracking"""
        self.print_status("Launching browser and starting scraping process...", "progress")
        
        browser = await launch(
            headless=True,
            executablePath="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
        )
        all_emails = []

        for i, url in enumerate(urls, 1):
            self.print_progress_bar(i, len(urls), prefix='Scraping Progress:', suffix=f'({i}/{len(urls)})')
            self.print_status(f"Processing: {url[:60]}...", "link")
            
            try:
                page = await browser.newPage()
                await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/91.0.864.48")
                await page.goto(url, {"waitUntil": "networkidle2"})

                # Search for "Contact" links
                contact_links = await page.evaluate('''() => {
                    return Array.from(document.querySelectorAll('a'))
                                .filter(a => a.textContent.toLowerCase().includes('contact'))
                                .map(a => a.href);
                }''')

                if contact_links:
                    self.print_status(f"Found {len(contact_links)} contact links", "success")
                    for contact_url in contact_links[:3]:  # Limit to 3 contact pages
                        try:
                            await page.goto(contact_url, {"waitUntil": "networkidle2"})
                        except Exception as e:
                            continue

                page_content = await page.content()
                emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', page_content)

                if emails:
                    self.print_status(f"Found {len(emails)} email addresses", "email")
                    all_emails.extend(emails)
                else:
                    self.print_status("No emails found on this page", "warning")
            except Exception as e:
                self.print_status(f"Error scraping {url}: {str(e)[:50]}...", "error")
            finally:
                await page.close()

        await browser.close()
        print()  # New line after progress bar
        return all_emails

    def save_emails_to_file(self):
        """Save collected emails to a text file with timestamp using custom path"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Use custom save path instead of desktop
        if not os.path.exists(self.custom_save_path):
            os.makedirs(self.custom_save_path)
            
        output_file_path = os.path.join(self.custom_save_path, f"collected_emails_{timestamp}.txt")
        
        with open(output_file_path, "w", encoding="utf-8") as file:
            file.write(f"# Email Collection Report\n")
            file.write(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            file.write(f"# Total Emails: {len(self.collected_emails)}\n")
            file.write(f"# {'='*50}\n\n")
            for email in self.collected_emails:
                file.write(f"{email}\n")
        
        return output_file_path

    def send_emails_flow(self):
        """Enhanced email sending process"""
        if not self.collected_emails:
            self.print_status("No emails available! Please collect or import emails first.", "error")
            self.press_enter_to_continue()
            return
        
        self.print_banner()
        self.print_header("EMAIL CAMPAIGN WIZARD")
        
        self.print_section("CAMPAIGN SETUP")
        self.print_status(f"Using {len(self.collected_emails)} emails for campaign", "success")
        
        # Email credentials
        sender_email = input(colored("📧 Enter your email address: ", "yellow")).strip()
        password = getpass(colored("🔒 Enter your email password: ", "yellow")).strip()
        
        # Email content
        self.print_section("EMAIL CONTENT")
        letter_path = input(colored("📝 Drag and drop your Letter.txt file here: ", "yellow")).strip().strip('"')
        
        try:
            with open(letter_path, "r", encoding="utf-8") as letter_file:
                subject = letter_file.readline().strip()
                body = letter_file.read().strip()
            self.print_status("Email content loaded successfully!", "success")
            self.print_status(f"Subject: {subject}", "info")
        except FileNotFoundError:
            self.print_status("Error: Letter file not found.", "error")
            self.press_enter_to_continue()
            return
        
        # Attachments
        self.print_section("ATTACHMENTS")
        cv_path = input(colored("📎 Drag and drop your CV file here: ", "yellow")).strip().strip('"')
        demand_path = input(colored("📎 Drag and drop your Demand file here: ", "yellow")).strip().strip('"')
        
        files_to_attach = []
        for path, name in [(cv_path, "CV"), (demand_path, "Demand Letter")]:
            if path and os.path.exists(path):
                files_to_attach.append(path)
                self.print_status(f"{name} attached: {os.path.basename(path)}", "success")
            else:
                self.print_status(f"{name} not found or invalid", "warning")
        
        # Confirmation
        self.print_section("CONFIRMATION")
        print(colored("🎯 CAMPAIGN SUMMARY:", "magenta"))
        print(colored(f"   • Recipients: {len(self.collected_emails)}", "cyan"))
        print(colored(f"   • Subject: {subject}", "cyan"))
        print(colored(f"   • Attachments: {len(files_to_attach)}", "cyan"))
        print(colored(f"   • Estimated time: {len(self.collected_emails) * 15} seconds", "cyan"))
        
        confirm = input(colored("\n🚀 Proceed with email campaign? (y/n): ", "red")).lower()
        if confirm != 'y':
            self.print_status("Email campaign cancelled.", "warning")
            self.press_enter_to_continue()
            return
        
        # Send emails
        self.send_emails(sender_email, password, self.collected_emails, subject, body, files_to_attach)
        self.press_enter_to_continue()

    def send_emails(self, sender_email, password, recipient_emails, subject, body, files_to_attach):
        """Enhanced email sending with detailed tracking"""
        self.print_banner()
        self.print_header("EMAIL CAMPAIGN IN PROGRESS")
        
        successful_sends = 0
        failed_sends = 0
        
        self.print_section("SENDING PROGRESS")
        
        for i, receiver_email in enumerate(recipient_emails, 1):
            # Progress display
            self.print_progress_bar(i, len(recipient_emails), 
                                 prefix='Sending Emails:', 
                                 suffix=f'({successful_sends}✅ {failed_sends}❌)')
            
            print(colored(f"\n  📧 Sending to: {receiver_email}", "blue"))
            
            try:
                # Create email
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = subject
                message.attach(MIMEText(body, "plain"))

                # Attach files
                for filepath in files_to_attach:
                    filename = os.path.basename(filepath)
                    try:
                        with open(filepath, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={filename}")
                        message.attach(part)
                    except FileNotFoundError:
                        self.print_status(f"Attachment {filename} not found", "warning")

                # Send email
                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                    server.starttls()
                    server.login(sender_email, password)
                    server.sendmail(sender_email, receiver_email, message.as_string())
                
                self.print_status("Email sent successfully!", "success")
                successful_sends += 1
                self.session_stats['successful_sends'] += 1
                
            except Exception as e:
                error_msg = str(e)
                self.print_status(f"Failed to send: {error_msg[:60]}...", "error")
                failed_sends += 1
                self.session_stats['failed_sends'] += 1

            # Delay between emails
            if i < len(recipient_emails):
                print(colored("  ⏳ Waiting 15 seconds before next email...", "blue"))
                for remaining in range(15, 0, -1):
                    print(f"\r  ⏰ Next email in: {remaining:2d} seconds", end='')
                    time.sleep(1)
                print("\r" + " " * 50, end='\r')  # Clear line
        
        # Final summary with fixed table formatting
        self.print_section("CAMPAIGN COMPLETE")
        print(colored("🎊 EMAIL CAMPAIGN SUMMARY:", "magenta"))
        print(colored("┌──────────────────────┬──────────┐", "green"))
        print(colored("│       Metric         │  Count   │", "green"))
        print(colored("├──────────────────────┼──────────┤", "green"))
        print(colored(f"│  ✅ Successful Sends  │ {successful_sends:8} │", "green"))
        print(colored(f"│  ❌ Failed Sends      │ {failed_sends:8} │", "red"))
        print(colored(f"│  📊 Total Attempted   │ {len(recipient_emails):8} │", "cyan"))
        print(colored("└──────────────────────┴──────────┘", "green"))
        
        success_rate = (successful_sends / len(recipient_emails)) * 100
        print(colored(f"📈 Success Rate: {success_rate:.1f}%", 
                    "green" if success_rate > 80 else "yellow" if success_rate > 50 else "red"))

    def view_collected_emails(self):
        """Enhanced email viewing interface"""
        self.print_banner()
        self.print_header("COLLECTED EMAILS DATABASE")
        
        if not self.collected_emails:
            self.print_status("No emails collected yet!", "warning")
            self.press_enter_to_continue()
            return
        
        self.print_section(f"EMAIL LIST - {len(self.collected_emails)} ADDRESSES")
        
        # Pagination or full list for small collections
        if len(self.collected_emails) <= 20:
            for i, email in enumerate(self.collected_emails, 1):
                print(colored(f"  {i:3d}. {email}", "cyan"))
        else:
            # Show first 10 and last 10
            for i in range(10):
                print(colored(f"  {i+1:3d}. {self.collected_emails[i]}", "cyan"))
            print(colored(f"  ... {len(self.collected_emails) - 20} more emails ...", "blue"))
            for i in range(-10, 0):
                print(colored(f"  {len(self.collected_emails)+i+1:3d}. {self.collected_emails[i]}", "cyan"))
        
        self.print_section_end()
        self.press_enter_to_continue()

    def clear_emails(self):
        """Enhanced email clearing with confirmation"""
        self.print_banner()
        self.print_header("DATA MANAGEMENT")
        
        if not self.collected_emails:
            self.print_status("No emails to clear!", "info")
            self.press_enter_to_continue()
            return
        
        self.print_section("CONFIRM DATA CLEARANCE")
        print(colored("⚠️  WARNING: This action cannot be undone!", "red"))
        print(colored(f"📧 You are about to clear {len(self.collected_emails)} collected emails", "yellow"))
        
        confirm = input(colored("\n🔴 Type 'CONFIRM' to proceed: ", "red")).strip()
        if confirm.upper() == "CONFIRM":
            self.collected_emails = []
            self.print_status("All collected emails have been cleared!", "success")
        else:
            self.print_status("Data clearance cancelled.", "info")
        
        self.press_enter_to_continue()

    def show_session_report(self):
        """Display detailed session statistics"""
        self.print_banner()
        self.print_header("SESSION ANALYTICS REPORT")
        
        duration = datetime.now() - self.session_stats['start_time'] if self.session_stats['start_time'] else timedelta(0)
        
        print(colored("\n📈 PERFORMANCE METRICS:", "magenta"))
        print(colored("┌──────────────────────────────┬────────────────────┐", "cyan"))
        metrics = [
            ("Session Duration", f"{duration}"),
            ("Emails Collected", f"{len(self.collected_emails)}"),
            ("Successful Sends", f"{self.session_stats['successful_sends']}"),
            ("Failed Sends", f"{self.session_stats['failed_sends']}"),
            ("Total Operations", f"{self.session_stats['successful_sends'] + self.session_stats['failed_sends']}"),
        ]
        
        for metric, value in metrics:
            print(colored(f"│ {metric:28} │ {value:18} │", "white"))
        print(colored("└──────────────────────────────┴────────────────────┘", "cyan"))
        
        if self.session_stats['successful_sends'] + self.session_stats['failed_sends'] > 0:
            success_rate = (self.session_stats['successful_sends'] / 
                          (self.session_stats['successful_sends'] + self.session_stats['failed_sends'])) * 100
            print(colored(f"\n🎯 Success Rate: {success_rate:.1f}%", 
                        "green" if success_rate > 80 else "yellow" if success_rate > 50 else "red"))
        
        self.press_enter_to_continue()

    def exit_gracefully(self):
        """Enhanced exit with session summary"""
        self.print_banner()
        self.print_header("SESSION SUMMARY")
        
        duration = datetime.now() - self.session_stats['start_time'] if self.session_stats['start_time'] else timedelta(0)
        
        print(colored("🎊 Thank you for using Ultimate Email Automation Suite!", "green"))
        print(colored(f"⏱️  Session duration: {duration}", "cyan"))
        print(colored(f"📧 Emails processed: {len(self.collected_emails)}", "cyan"))
        print(colored(f"🚀 Emails sent: {self.session_stats['successful_sends']}", "cyan"))
        
        print(colored("\n👋 Goodbye!", "yellow"))
        sys.exit(0)

    def press_enter_to_continue(self):
        """Uniform continue prompt"""
        input(colored("\n↵ Press Enter to continue...", "blue"))

def main():
    """Main function to run the ultimate email automation suite"""
    try:
        automation_suite = UltimateEmailAutomationSuite()
        automation_suite.display_menu()
    except KeyboardInterrupt:
        print(colored("\n\n🛑 Program interrupted by user.", "red"))
    except Exception as e:
        print(colored(f"\n💥 An unexpected error occurred: {e}", "red"))

if __name__ == "__main__":
    main()