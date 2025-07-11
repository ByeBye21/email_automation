"""
Email Automation Application
Version 1.0.0

A comprehensive email automation tool with CSV recipients, template personalization,
SMTP sending, and professional GUI interface.

Author: ByeBye21
Email: younes0079@gmail.com
GitHub: https://github.com/ByeBye21
Copyright © 2025 ByeBye21. All rights reserved.

Date: July 10, 2025
"""

import os
import sys
import argparse
import json
import csv
import logging
import smtplib
import ssl
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Union
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email import encoders
import mimetypes
import re

# Third-party imports
try:
    from jinja2 import Template, Environment, FileSystemLoader, TemplateNotFound
    JINJA2_AVAILABLE = True
except ImportError:
    print("Warning: Jinja2 not available. Template functionality will be limited.")
    JINJA2_AVAILABLE = False
    Template = None
    Environment = None
    FileSystemLoader = None
    TemplateNotFound = Exception


class EmailAppError(Exception):
    """Base exception for email application errors."""
    pass


class ConfigurationError(EmailAppError):
    """Exception for configuration-related errors."""
    pass


class EmailSendError(EmailAppError):
    """Exception for email sending errors."""
    pass


class TemplateError(EmailAppError):
    """Exception for template-related errors."""
    pass


class CSVError(EmailAppError):
    """Exception for CSV-related errors."""
    pass


class Logger:
    """
    Centralized logging system for the email application.
    
    Provides structured logging with file and console output,
    specific methods for different types of operations.
    """
    
    def __init__(self, log_file: str = "email_app.log", log_level: str = "INFO"):
        """
        Initialize the logger.
        
        Args:
            log_file (str): Path to the log file
            log_level (str): Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        """
        self.log_file = log_file
        self.log_level = getattr(logging, log_level.upper(), logging.INFO)
        self.logger = logging.getLogger('email_app')
        self._setup_logging()
    
    def _setup_logging(self) -> None:
        """Configure logging with both file and console handlers."""
        # Clear existing handlers
        self.logger.handlers.clear()
        self.logger.setLevel(self.log_level)
        
        # Create formatters
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        simple_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # File handler
        try:
            log_path = Path(self.log_file)
            log_path.parent.mkdir(parents=True, exist_ok=True)
            
            file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
            file_handler.setLevel(self.log_level)
            file_handler.setFormatter(detailed_formatter)
            self.logger.addHandler(file_handler)
        except Exception as e:
            print(f"Warning: Could not create log file {self.log_file}: {e}")
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(simple_formatter)
        self.logger.addHandler(console_handler)
        
        self.info(f"Logging initialized - File: {self.log_file}")
    
    def debug(self, message: str) -> None:
        """Log debug message."""
        self.logger.debug(message)
    
    def info(self, message: str) -> None:
        """Log info message."""
        self.logger.info(message)
    
    def warning(self, message: str) -> None:
        """Log warning message."""
        self.logger.warning(message)
    
    def error(self, message: str) -> None:
        """Log error message."""
        self.logger.error(message)
    
    def critical(self, message: str) -> None:
        """Log critical message."""
        self.logger.critical(message)
    
    def log_email_success(self, recipient: str, subject: str, attachments: int = 0) -> None:
        """Log successful email send."""
        attachment_info = f" with {attachments} attachment(s)" if attachments else ""
        self.info(f"EMAIL SENT ✓ - To: {recipient} - Subject: '{subject}'{attachment_info}")
    
    def log_email_failure(self, recipient: str, error: str, subject: str = None) -> None:
        """Log failed email send."""
        subject_info = f" - Subject: '{subject}'" if subject else ""
        self.error(f"EMAIL FAILED ✗ - To: {recipient}{subject_info} - Error: {error}")
    
    def log_smtp_connection(self, host: str, port: int, username: str) -> None:
        """Log SMTP connection attempt."""
        self.info(f"SMTP CONNECT - {host}:{port} - User: {username}")
    
    def log_test_result(self, success: bool, recipient: str, error: str = None) -> None:
        """Log test email result."""
        if success:
            self.info(f"TEST EMAIL ✓ - Sent to: {recipient}")
        else:
            self.error(f"TEST EMAIL ✗ - Failed to: {recipient} - Error: {error}")


class ConfigurationManager:
    """
    Manages application configuration with validation and default values.
    
    Loads JSON configuration files and provides structured access to
    SMTP settings, file paths, email settings, and application options.
    """
    
    # Default configuration structure
    DEFAULT_CONFIG = {
        'smtp': {
            'server': '',
            'port': 587,
            'username': '',
            'password': '',
            'use_tls': True,
            'use_ssl': False,
            'timeout': 30
        },
        'files': {
            'csv_recipients': 'recipients.csv',
            'email_template': 'template.txt',
            'output_directory': './output',
            'log_directory': './logs'
        },
        'email': {
            'default_subject': 'Email from {{sender_name}}',
            'sender_name': 'Email Application',
            'test_email': None,
            'reply_to': None
        },
        'application': {
            'debug': False,
            'max_attachments': 10,
            'max_file_size_mb': 25,
            'batch_size': 100,
            'template_directory': './templates'
        }
    }
    
    def __init__(self, logger: Logger):
        """
        Initialize configuration manager.
        
        Args:
            logger (Logger): Application logger instance
        """
        self.logger = logger
        self.config = {}
        self.is_loaded = False
    
    def load_config(self, config_file: str) -> Dict[str, Any]:
        """
        Load and validate configuration from JSON file.
        
        Args:
            config_file (str): Path to configuration file
            
        Returns:
            Dict[str, Any]: Loaded configuration
            
        Raises:
            ConfigurationError: If configuration is invalid
        """
        if not os.path.exists(config_file):
            raise ConfigurationError(f"Configuration file not found: {config_file}")
        
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
        except json.JSONDecodeError as e:
            raise ConfigurationError(f"Invalid JSON in {config_file}: {e}")
        except Exception as e:
            raise ConfigurationError(f"Error reading {config_file}: {e}")
        
        # Merge with defaults
        self.config = self._merge_config(self.DEFAULT_CONFIG, user_config)
        
        # Validate configuration
        self._validate_config()
        
        self.is_loaded = True
        self.logger.info(f"Configuration loaded from {config_file}")
        return self.config.copy()
    
    def _merge_config(self, default: Dict[str, Any], user: Dict[str, Any]) -> Dict[str, Any]:
        """Merge user configuration with defaults."""
        result = default.copy()
        
        for key, value in user.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_config(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def _validate_config(self) -> None:
        """Validate the loaded configuration."""
        errors = []
        
        # Validate SMTP settings
        smtp = self.config.get('smtp', {})
        required_smtp = ['server', 'username', 'password']
        for field in required_smtp:
            if not smtp.get(field):
                errors.append(f"Missing required SMTP field: {field}")
        
        # Validate port
        port = smtp.get('port')
        if not isinstance(port, int) or not (1 <= port <= 65535):
            errors.append("SMTP port must be an integer between 1 and 65535")
        
        # Validate file paths
        files = self.config.get('files', {})
        if not files.get('csv_recipients'):
            errors.append("CSV recipients file path is required")
        if not files.get('email_template'):
            errors.append("Email template file path is required")
        
        if errors:
            raise ConfigurationError("Configuration validation failed:\n" + "\n".join(errors))
    
    def get(self, section: str, key: str = None, default: Any = None) -> Any:
        """Get configuration value."""
        if not self.is_loaded:
            raise ConfigurationError("Configuration not loaded")
        
        if key is None:
            return self.config.get(section, default)
        return self.config.get(section, {}).get(key, default)


class CSVReader:
    """
    Handles reading and processing CSV recipient data.
    
    Provides functionality to read CSV files with validation,
    error handling, and data cleaning.
    """
    
    def __init__(self, logger: Logger):
        """
        Initialize CSV reader.
        
        Args:
            logger (Logger): Application logger instance
        """
        self.logger = logger
    
    def read_recipients(self, csv_file: str) -> List[Dict[str, Any]]:
        """
        Read recipient data from CSV file.
        
        Args:
            csv_file (str): Path to CSV file
            
        Returns:
            List[Dict[str, Any]]: List of recipient data dictionaries
            
        Raises:
            CSVError: If CSV reading fails
        """
        if not os.path.exists(csv_file):
            raise CSVError(f"CSV file not found: {csv_file}")
        
        recipients = []
        
        try:
            with open(csv_file, 'r', newline='', encoding='utf-8') as csvfile:
                # Auto-detect CSV format
                sample = csvfile.read(1024)
                csvfile.seek(0)
                
                try:
                    dialect = csv.Sniffer().sniff(sample)
                    delimiter = dialect.delimiter
                except csv.Error:
                    delimiter = ','
                
                reader = csv.DictReader(csvfile, delimiter=delimiter)
                
                if not reader.fieldnames:
                    raise CSVError("CSV file appears to be empty or has no headers")
                
                # Validate required columns
                self.logger.info(f"CSV columns found: {', '.join(reader.fieldnames)}")
                
                for row_num, row in enumerate(reader, start=2):
                    # Skip empty rows
                    if not any(row.values()):
                        continue
                    
                    # Clean and process row data
                    cleaned_row = {}
                    for key, value in row.items():
                        cleaned_row[key] = str(value).strip() if value else ''
                    
                    recipients.append(cleaned_row)
                
                if not recipients:
                    raise CSVError("No valid recipients found in CSV file")
                
                self.logger.info(f"Loaded {len(recipients)} valid recipients from CSV")
                return recipients
                
        except csv.Error as e:
            raise CSVError(f"CSV parsing error: {e}")
        except UnicodeDecodeError:
            raise CSVError("CSV file encoding error. Please ensure UTF-8 encoding.")
        except Exception as e:
            raise CSVError(f"Error reading CSV file: {e}")


class TemplateRenderer:
    """
    Handles email template rendering using Jinja2 or basic string replacement.
    
    Provides personalized email content generation with support for
    both Jinja2 templates and simple placeholder replacement.
    """
    
    def __init__(self, logger: Logger, template_directory: str = None):
        """
        Initialize template renderer.
        
        Args:
            logger (Logger): Application logger instance
            template_directory (str): Directory containing templates
        """
        self.logger = logger
        self.template_directory = template_directory
        self.jinja_env = None
        
        if JINJA2_AVAILABLE and template_directory:
            self.jinja_env = Environment(loader=FileSystemLoader(template_directory))
    
    def render_template(self, template_path: str, data: Dict[str, Any]) -> str:
        """
        Render template with recipient data.
        
        Args:
            template_path (str): Path to template file
            data (Dict[str, Any]): Data for template rendering
            
        Returns:
            str: Rendered template content
            
        Raises:
            TemplateError: If template rendering fails
        """
        if not os.path.exists(template_path):
            raise TemplateError(f"Template file not found: {template_path}")
        
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_content = f.read()
        except Exception as e:
            raise TemplateError(f"Error reading template file: {e}")
        
        try:
            if JINJA2_AVAILABLE:
                # Use Jinja2 for advanced templating
                template = Template(template_content)
                rendered = template.render(**data)
            else:
                # Use basic string replacement
                rendered = self._basic_template_replace(template_content, data)
            
            self.logger.debug(f"Template rendered for {data.get('email', 'unknown recipient')}")
            return rendered
            
        except Exception as e:
            raise TemplateError(f"Template rendering error: {e}")
    
    def _basic_template_replace(self, content: str, data: Dict[str, Any]) -> str:
        """Basic template replacement for when Jinja2 is not available."""
        rendered = content
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            rendered = rendered.replace(placeholder, str(value))
        return rendered


class AttachmentHandler:
    """
    Handles email attachments with MIME type detection and validation.
    
    Creates appropriate MIME objects for different file types
    with size limits and error handling.
    """
    
    def __init__(self, logger: Logger, max_size_mb: int = 25):
        """
        Initialize attachment handler.
        
        Args:
            logger (Logger): Application logger instance
            max_size_mb (int): Maximum file size in megabytes
        """
        self.logger = logger
        self.max_size_bytes = max_size_mb * 1024 * 1024
    
    def create_attachment(self, file_path: str) -> MIMEBase:
        """
        Create email attachment from file.
        
        Args:
            file_path (str): Path to file to attach
            
        Returns:
            MIMEBase: Email attachment object
            
        Raises:
            EmailSendError: If attachment creation fails
        """
        if not os.path.exists(file_path):
            raise EmailSendError(f"Attachment file not found: {file_path}")
        
        if not os.path.isfile(file_path):
            raise EmailSendError(f"Attachment path is not a file: {file_path}")
        
        file_size = os.path.getsize(file_path)
        if file_size > self.max_size_bytes:
            size_mb = file_size / (1024 * 1024)
            max_mb = self.max_size_bytes / (1024 * 1024)
            raise EmailSendError(f"File too large: {size_mb:.1f}MB (max: {max_mb}MB)")
        
        file_name = os.path.basename(file_path)
        content_type, encoding = mimetypes.guess_type(file_path)
        
        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        
        main_type, sub_type = content_type.split('/', 1)
        
        try:
            with open(file_path, 'rb') as f:
                file_data = f.read()
        except Exception as e:
            raise EmailSendError(f"Cannot read attachment file: {e}")
        
        # Create appropriate MIME object
        if main_type == 'text':
            attachment = MIMEText(file_data.decode('utf-8', errors='replace'), _subtype=sub_type)
        elif main_type == 'image':
            attachment = MIMEImage(file_data, _subtype=sub_type)
        elif main_type == 'audio':
            attachment = MIMEAudio(file_data, _subtype=sub_type)
        else:
            attachment = MIMEApplication(file_data, _subtype=sub_type)
        
        attachment.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
        
        self.logger.debug(f"Created attachment: {file_name} ({file_size} bytes)")
        return attachment


class EmailSender:
    """
    Handles SMTP email sending with comprehensive error handling and logging.
    
    Supports TLS/SSL connections, multiple recipients, HTML/text content,
    and attachments with detailed logging of all operations.
    """
    
    def __init__(self, logger: Logger, attachment_handler: AttachmentHandler):
        """
        Initialize email sender.
        
        Args:
            logger (Logger): Application logger instance
            attachment_handler (AttachmentHandler): Attachment handler instance
        """
        self.logger = logger
        self.attachment_handler = attachment_handler
    
    def send_email(self, 
                   smtp_config: Dict[str, Any],
                   sender: str,
                   recipient: Union[str, List[str]],
                   subject: str,
                   body: str,
                   html_body: str = None,
                   attachments: List[str] = None,
                   sender_name: str = None) -> bool:
        """
        Send email via SMTP.
        
        Args:
            smtp_config (Dict[str, Any]): SMTP configuration
            sender (str): Sender email address
            recipient (Union[str, List[str]]): Recipient email address(es)
            subject (str): Email subject
            body (str): Email body (plain text)
            html_body (str, optional): HTML email body
            attachments (List[str], optional): List of attachment file paths
            sender_name (str, optional): Sender display name
            
        Returns:
            bool: True if email sent successfully
            
        Raises:
            EmailSendError: If email sending fails
        """
        # Normalize recipients
        recipients = [recipient] if isinstance(recipient, str) else recipient
        if not recipients:
            raise EmailSendError("No recipients specified")
        
        recipient_str = ', '.join(recipients)
        self.logger.info(f"Sending email to: {recipient_str}")
        self.logger.log_smtp_connection(
            smtp_config['server'], 
            smtp_config['port'], 
            smtp_config['username']
        )
        
        # Create message
        msg = MIMEMultipart('alternative')
        msg['From'] = f"{sender_name} <{sender}>" if sender_name else sender
        msg['To'] = recipient_str
        msg['Subject'] = subject
        
        # Add body content
        if html_body:
            text_part = MIMEText(body, 'plain')
            html_part = MIMEText(html_body, 'html')
            msg.attach(text_part)
            msg.attach(html_part)
        else:
            text_part = MIMEText(body, 'plain')
            msg.attach(text_part)
        
        # Add attachments
        attachment_count = 0
        if attachments:
            for file_path in attachments:
                try:
                    attachment = self.attachment_handler.create_attachment(file_path)
                    msg.attach(attachment)
                    attachment_count += 1
                except Exception as e:
                    self.logger.error(f"Failed to attach {file_path}: {e}")
                    raise EmailSendError(f"Attachment error: {e}")
        
        # Send email
        try:
            with self._create_smtp_connection(smtp_config) as smtp_server:
                smtp_server.send_message(msg)
                
                for recipient_email in recipients:
                    self.logger.log_email_success(recipient_email, subject, attachment_count)
                
                return True
                
        except Exception as e:
            error_msg = str(e)
            for recipient_email in recipients:
                self.logger.log_email_failure(recipient_email, error_msg, subject)
            raise EmailSendError(f"SMTP error: {error_msg}")
    
    def _create_smtp_connection(self, smtp_config: Dict[str, Any]):
        """Create and configure SMTP connection."""
        try:
            if smtp_config.get('use_ssl', False):
                context = ssl.create_default_context()
                smtp = smtplib.SMTP_SSL(
                    smtp_config['server'],
                    smtp_config['port'],
                    context=context,
                    timeout=smtp_config.get('timeout', 30)
                )
            else:
                smtp = smtplib.SMTP(
                    smtp_config['server'],
                    smtp_config['port'],
                    timeout=smtp_config.get('timeout', 30)
                )
                
                if smtp_config.get('use_tls', True):
                    smtp.starttls()
            
            smtp.login(smtp_config['username'], smtp_config['password'])
            return smtp
            
        except smtplib.SMTPAuthenticationError as e:
            raise EmailSendError(f"SMTP authentication failed: {e}")
        except smtplib.SMTPException as e:
            raise EmailSendError(f"SMTP error: {e}")
        except Exception as e:
            raise EmailSendError(f"Connection error: {e}")


class EmailApplication:
    """
    Main application class that orchestrates all components.
    
    Integrates configuration management, CSV reading, template rendering,
    and email sending into a cohesive application with test functionality.
    """
    
    def __init__(self, config_file: str = None, log_file: str = None, log_level: str = "INFO"):
        """
        Initialize the email application.
        
        Args:
            config_file (str, optional): Path to configuration file
            log_file (str, optional): Path to log file
            log_level (str): Logging level
        """
        # Initialize logger
        log_path = log_file or "email_app.log"
        self.logger = Logger(log_path, log_level)
        
        # Initialize components
        self.config_manager = ConfigurationManager(self.logger)
        self.csv_reader = CSVReader(self.logger)
        self.attachment_handler = AttachmentHandler(self.logger)
        self.email_sender = EmailSender(self.logger, self.attachment_handler)
        self.template_renderer = None
        
        # Load configuration if provided
        if config_file:
            self.load_config(config_file)
    
    def load_config(self, config_file: str) -> None:
        """Load application configuration."""
        self.config = self.config_manager.load_config(config_file)
        
        # Initialize template renderer with config
        template_dir = self.config.get('application', 'template_directory')
        self.template_renderer = TemplateRenderer(self.logger, template_dir)
        
        # Update attachment handler max size
        max_size = self.config.get('application', 'max_file_size_mb', 25)
        self.attachment_handler.max_size_bytes = max_size * 1024 * 1024
    
    def send_test_email(self, test_recipient: str = None) -> bool:
        """
        Send test email to verify configuration.
        
        Args:
            test_recipient (str, optional): Test email recipient
            
        Returns:
            bool: True if test successful
        """
        if not self.config:
            raise EmailAppError("No configuration loaded")
        
        smtp_config = self.config['smtp']
        email_config = self.config['email']
        
        # Determine test recipient
        recipient = (test_recipient or 
                    email_config.get('test_email') or 
                    smtp_config['username'])
        
        # Create test email content
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        subject = f"Email Application Test - {timestamp}"
        
        body = f"""
Email Application Configuration Test

This test email verifies that your SMTP configuration is working correctly.

Configuration Details:
- SMTP Server: {smtp_config['server']}:{smtp_config['port']}
- Username: {smtp_config['username']}
- TLS Enabled: {smtp_config.get('use_tls', False)}
- SSL Enabled: {smtp_config.get('use_ssl', False)}
- Timestamp: {timestamp}

If you receive this email, your configuration is working properly!

Best regards,
Email Application
        """.strip()
        
        try:
            success = self.email_sender.send_email(
                smtp_config=smtp_config,
                sender=smtp_config['username'],
                recipient=recipient,
                subject=subject,
                body=body,
                sender_name=email_config.get('sender_name', 'Email Application Test')
            )
            
            self.logger.log_test_result(success, recipient)
            return success
            
        except Exception as e:
            self.logger.log_test_result(False, recipient, str(e))
            return False
    
    def send_bulk_emails(self, 
                        csv_file: str = None, 
                        template_file: str = None,
                        subject: str = None,
                        attachments: List[str] = None) -> Dict[str, Any]:
        """
        Send bulk emails to recipients from CSV.
        
        Args:
            csv_file (str, optional): Path to CSV file (uses config if None)
            template_file (str, optional): Path to template file (uses config if None)
            subject (str, optional): Email subject (uses config if None)
            attachments (List[str], optional): List of attachment paths
            
        Returns:
            Dict[str, Any]: Results summary
        """
        if not self.config:
            raise EmailAppError("No configuration loaded")
        
        # Use config values as defaults
        csv_path = csv_file or self.config['files']['csv_recipients']
        template_path = template_file or self.config['files']['email_template']
        email_subject = subject or self.config['email']['default_subject']
        
        # Read recipients
        self.logger.info("Reading recipient data from CSV...")
        recipients = self.csv_reader.read_recipients(csv_path)
        
        # Initialize results tracking
        results = {
            'total': len(recipients),
            'sent': 0,
            'failed': 0,
            'errors': [],
            'success': False
        }
        
        smtp_config = self.config['smtp']
        email_config = self.config['email']
        
        self.logger.info(f"Starting bulk email send to {results['total']} recipients")
        
        # Send emails
        for i, recipient_data in enumerate(recipients, 1):
            try:
                # Render template with recipient data
                personalized_body = self.template_renderer.render_template(
                    template_path, 
                    recipient_data
                )
                
                # Render subject with recipient data
                if JINJA2_AVAILABLE:
                    subject_template = Template(email_subject)
                    personalized_subject = subject_template.render(**recipient_data)
                else:
                    personalized_subject = self.template_renderer._basic_template_replace(
                        email_subject, 
                        recipient_data
                    )
                
                # Send email
                success = self.email_sender.send_email(
                    smtp_config=smtp_config,
                    sender=smtp_config['username'],
                    recipient=recipient_data['email'],
                    subject=personalized_subject,
                    body=personalized_body,
                    attachments=attachments,
                    sender_name=email_config.get('sender_name')
                )
                
                if success:
                    results['sent'] += 1
                else:
                    results['failed'] += 1
                    results['errors'].append(f"Failed to send to {recipient_data['email']}")
                
                # Progress logging
                if i % 10 == 0:
                    self.logger.info(f"Progress: {i}/{results['total']} emails processed")
                
            except Exception as e:
                results['failed'] += 1
                error_msg = f"Error sending to {recipient_data.get('email', 'unknown')}: {str(e)}"
                results['errors'].append(error_msg)
                self.logger.error(error_msg)
        
        # Final results
        results['success'] = results['failed'] == 0
        
        self.logger.info(f"Bulk email completed - Sent: {results['sent']}, Failed: {results['failed']}")
        
        return results


def create_sample_config(output_file: str = "config.json") -> None:
    """Create a sample configuration file."""
    sample_config = {
        "_comment": "Email Application Configuration File",
        "_description": "Configure SMTP settings, file paths, and application options",
        
        "smtp": {
            "_comment": "SMTP server configuration",
            "server": "smtp.gmail.com",
            "port": 587,
            "username": "your_email@gmail.com",
            "password": "your_app_password",
            "use_tls": True,
            "use_ssl": False,
            "timeout": 30
        },
        
        "files": {
            "_comment": "File paths for input and output",
            "csv_recipients": "./recipients.csv",
            "email_template": "./template.txt",
            "output_directory": "./output",
            "log_directory": "./logs"
        },
        
        "email": {
            "_comment": "Email settings and defaults",
            "default_subject": "Newsletter - {{month}} {{year}}",
            "sender_name": "Newsletter Team",
            "test_email": "test@example.com",
            "reply_to": None
        },
        
        "application": {
            "_comment": "Application settings",
            "debug": False,
            "max_attachments": 10,
            "max_file_size_mb": 25,
            "batch_size": 100,
            "template_directory": "./templates"
        }
    }
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(sample_config, f, indent=2, ensure_ascii=False)
        print(f"Sample configuration created: {output_file}")
    except Exception as e:
        print(f"Error creating sample config: {e}")


def create_sample_files() -> None:
    """Create sample CSV and template files."""
    # Sample CSV
    csv_content = '''email,name,company,position
john.doe@example.com,John Doe,Acme Corp,Software Engineer
jane.smith@example.com,Jane Smith,Tech Solutions,Marketing Manager
bob.wilson@example.com,Bob Wilson,StartupXYZ,CEO
alice.brown@example.com,Alice Brown,Global Inc,Data Scientist
'''
    
    # Sample template
    template_content = '''Subject: Welcome {{name}}!

Dear {{name}},

Thank you for your interest in our services. We're excited to work with {{company}}!

{% if position %}
As a {{position}}, you have unique needs that our team can address.
{% endif %}

Best regards,
The Team

---
Email: {{email}}
Company: {{company}}
'''
    
    try:
        with open('recipients.csv', 'w', encoding='utf-8') as f:
            f.write(csv_content)
        print("Sample CSV created: recipients.csv")
        
        with open('template.txt', 'w', encoding='utf-8') as f:
            f.write(template_content)
        print("Sample template created: template.txt")
        
    except Exception as e:
        print(f"Error creating sample files: {e}")


def main():
    """Command-line interface for the email application."""
    parser = argparse.ArgumentParser(
        description='Complete Email Sender Application',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --test                          # Send test email
  %(prog)s --send                          # Send bulk emails
  %(prog)s --config custom.json --test     # Test with custom config
  %(prog)s --create-samples                # Create sample files
  %(prog)s --send --log-level DEBUG        # Send with debug logging
        """
    )
    
    parser.add_argument('--config', '-c',
                       default='config.json',
                       help='Configuration file path (default: config.json)')
    
    parser.add_argument('--test', '-t',
                       action='store_true',
                       help='Send test email to verify SMTP configuration')
    
    parser.add_argument('--test-recipient',
                       help='Email address for test email (optional)')
    
    parser.add_argument('--send', '-s',
                       action='store_true',
                       help='Send bulk emails from CSV')
    
    parser.add_argument('--csv-file',
                       help='CSV file path (overrides config)')
    
    parser.add_argument('--template-file',
                       help='Template file path (overrides config)')
    
    parser.add_argument('--subject',
                       help='Email subject (overrides config)')
    
    parser.add_argument('--attachments',
                       nargs='*',
                       help='Attachment file paths')
    
    parser.add_argument('--log-file',
                       help='Log file path (default: email_app.log)')
    
    parser.add_argument('--log-level',
                       choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                       default='INFO',
                       help='Logging level (default: INFO)')
    
    parser.add_argument('--create-samples',
                       action='store_true',
                       help='Create sample configuration and template files')
    
    parser.add_argument('--create-config',
                       help='Create sample configuration file with specified name')
    
    args = parser.parse_args()
    
    # Handle sample creation
    if args.create_samples:
        create_sample_config()
        create_sample_files()
        return 0
    
    if args.create_config:
        create_sample_config(args.create_config)
        return 0
    
    # Initialize application
    try:
        app = EmailApplication(
            config_file=args.config,
            log_file=args.log_file,
            log_level=args.log_level
        )
        
        if args.test:
            # Send test email
            print("Sending test email...")
            success = app.send_test_email(args.test_recipient)
            
            if success:
                print("✓ Test email sent successfully!")
                return 0
            else:
                print("✗ Test email failed. Check logs for details.")
                return 1
        
        elif args.send:
            # Send bulk emails
            print("Sending bulk emails...")
            
            results = app.send_bulk_emails(
                csv_file=args.csv_file,
                template_file=args.template_file,
                subject=args.subject,
                attachments=args.attachments
            )
            
            print(f"\nBulk Email Results:")
            print(f"  Total Recipients: {results['total']}")
            print(f"  Successfully Sent: {results['sent']}")
            print(f"  Failed: {results['failed']}")
            print(f"  Overall Success: {'✓' if results['success'] else '✗'}")
            
            if results['errors']:
                print(f"\nErrors:")
                for error in results['errors'][:5]:  # Show first 5 errors
                    print(f"  - {error}")
                if len(results['errors']) > 5:
                    print(f"  ... and {len(results['errors']) - 5} more errors")
            
            return 0 if results['success'] else 1
        
        else:
            # Show help if no action specified
            parser.print_help()
            return 0
            
    except EmailAppError as e:
        print(f"Application Error: {e}")
        return 1
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        return 1
    except Exception as e:
        print(f"Unexpected Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
