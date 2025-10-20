#!/usr/bin/env python3
"""
Calendar Invite Plugin
Sends calendar invites via email using iCalendar format.
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import os
import logging

logger = logging.getLogger(__name__)

class CalendarInvitePlugin:
    """Plugin for sending calendar invites via email."""
    
    def __init__(self):
        self.name = "calendar_invite"
        self.description = "Sends calendar invites via email"
        self.version = "1.0.0"
        
    async def execute(self, context):
        """Execute the calendar invite functionality."""
        try:
            event_data = context.get('event_data', {})
            recipient_email = event_data.get('recipient_email')
            subject = event_data.get('subject', 'Calendar Invite')
            start_time = event_data.get('start_time', datetime.now() + timedelta(hours=1))
            if isinstance(start_time, str):
                start_time = datetime.fromisoformat(start_time)
            end_time = event_data.get('end_time', start_time + timedelta(hours=1))
            if isinstance(end_time, str):
                end_time = datetime.fromisoformat(end_time)
            location = event_data.get('location', '')
            description = event_data.get('description', '')
            
            if not recipient_email:
                return {"success": False, "error": "Recipient email is required"}
            
            # Create calendar invite
            ics_content = self._create_ics_content(subject, start_time, end_time, location, description)
            
            # Send email with calendar invite
            success = await self._send_calendar_invite(recipient_email, subject, ics_content)
            
            if success:
                return {"success": True, "message": f"Calendar invite sent to {recipient_email}"}
            else:
                return {"success": False, "error": "Failed to send calendar invite"}
                
        except Exception as e:
            logger.error(f"Error in calendar invite plugin: {e}")
            return {"success": False, "error": str(e)}
    
    def _create_ics_content(self, subject, start_time, end_time, location, description):
        """Create iCalendar content for the event."""
        uid = f"{int(datetime.now().timestamp())}@jarvis-demo.com"
        
        ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Jarvis AI//Calendar Plugin//EN
CALSCALE:GREGORIAN
METHOD:REQUEST
BEGIN:VEVENT
UID:{uid}
DTSTAMP:{self._format_datetime(datetime.now())}
DTSTART:{self._format_datetime(start_time)}
DTEND:{self._format_datetime(end_time)}
SUMMARY:{subject}
LOCATION:{location}
DESCRIPTION:{description}
STATUS:CONFIRMED
SEQUENCE:0
BEGIN:VALARM
TRIGGER:-PT15M
DESCRIPTION:Reminder
ACTION:DISPLAY
END:VALARM
END:VEVENT
END:VCALENDAR"""
        
        return ics_content
    
    def _format_datetime(self, dt):
        """Format datetime to iCalendar format (YYYYMMDDTHHMMSSZ)."""
        return dt.strftime('%Y%m%dT%H%M%SZ')
    
    async def _send_calendar_invite(self, recipient_email, subject, ics_content):
        """Send calendar invite via email."""
        try:
            # Get SMTP settings from environment
            smtp_host = os.getenv('SMTP_HOST', 'smtp.gmail.com')
            smtp_port = int(os.getenv('SMTP_PORT', '587'))
            smtp_user = os.getenv('SMTP_USER', '')
            smtp_pass = os.getenv('SMTP_PASS', '')
            from_email = os.getenv('FROM_EMAIL', smtp_user)
            
            if not smtp_user or not smtp_pass:
                logger.error("SMTP credentials not configured")
                return False
            
            # Create email message
            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = recipient_email
            msg['Subject'] = f"Invitation: {subject}"
            
            # Add text body
            text_body = f"You have been invited to: {subject}\n\nPlease see the attached calendar invite."
            msg.attach(MIMEText(text_body, 'plain'))
            
            # Add calendar attachment
            calendar_part = MIMEText(ics_content, 'calendar;method=REQUEST')
            calendar_part.add_header('Content-Disposition', 'attachment; filename="invite.ics"')
            msg.attach(calendar_part)
            
            # Send email
            with smtplib.SMTP(smtp_host, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
            
            logger.info(f"Calendar invite sent to {recipient_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send calendar invite: {e}")
            return False

# Plugin registration
def get_plugin():
    """Return the plugin instance."""
    return CalendarInvitePlugin()
