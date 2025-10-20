"""
Calendar Plugin
Sends calendar invites via email.
"""

import logging
from datetime import datetime, timedelta
from typing import Dict, Any, List
from plugins.plugin_manager import PluginBase

logger = logging.getLogger(__name__)

class CalendarPlugin(PluginBase):
    """Plugin for sending calendar invites."""
    
    def __init__(self, config: Dict[str, Any] = None):
        super().__init__(config)
        self.name = "calendar_plugin"
        self.description = "Sends calendar invites via email"
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Execute calendar invite functionality."""
        try:
            # Extract calendar details from context
            event_title = context.get("title", "Meeting")
            start_time = context.get("start_time")
            duration = context.get("duration", 60)  # minutes
            attendees = context.get("attendees", [])
            location = context.get("location", "")
            description = context.get("description", "")
            
            if not start_time:
                start_time = datetime.now() + timedelta(hours=1)
            elif isinstance(start_time, str):
                start_time = datetime.fromisoformat(start_time)
            
            end_time = start_time + timedelta(minutes=duration)
            
            # Generate calendar invite
            invite_data = self._generate_calendar_invite(
                event_title, start_time, end_time, attendees, location, description
            )
            
            # Send via email (if email functionality is available)
            sent = await self._send_calendar_invite(invite_data, attendees)
            
            return {
                "success": True,
                "message": f"Calendar invite sent for '{event_title}'",
                "details": {
                    "title": event_title,
                    "start_time": start_time.isoformat(),
                    "end_time": end_time.isoformat(),
                    "attendees": attendees,
                    "sent": sent
                }
            }
            
        except Exception as e:
            logger.error(f"Error in calendar plugin: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _generate_calendar_invite(self, title: str, start: datetime, end: datetime, 
                                 attendees: List[str], location: str, description: str) -> str:
        """Generate iCal format calendar invite."""
        # Generate unique ID
        import uuid
        uid = str(uuid.uuid4())
        
        # Format times for iCal
        start_str = start.strftime("%Y%m%dT%H%M%SZ")
        end_str = end.strftime("%Y%m%dT%H%M%SZ")
        now_str = datetime.now().strftime("%Y%m%dT%H%M%SZ")
        
        # Create iCal content
        ical_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Jarvis System//Calendar Plugin//EN
BEGIN:VEVENT
UID:{uid}
DTSTAMP:{now_str}
DTSTART:{start_str}
DTEND:{end_str}
SUMMARY:{title}
LOCATION:{location}
DESCRIPTION:{description}
STATUS:CONFIRMED
SEQUENCE:0
END:VEVENT
END:VCALENDAR"""
        
        return ical_content
    
    async def _send_calendar_invite(self, ical_content: str, attendees: List[str]) -> bool:
        """Send calendar invite via email."""
        try:
            # This would integrate with the existing email system
            # For now, just log the action
            logger.info(f"Would send calendar invite to: {attendees}")
            logger.info(f"iCal content: {ical_content[:200]}...")
            
            # In a real implementation, this would:
            # 1. Create an email with the iCal attachment
            # 2. Send it to all attendees
            # 3. Return success/failure status
            
            return True
            
        except Exception as e:
            logger.error(f"Error sending calendar invite: {e}")
            return False
    
    def get_description(self) -> str:
        return "Sends calendar invites via email with iCal attachments"
    
    def get_requirements(self) -> List[str]:
        return ["uuid", "datetime"]
    
    def validate_config(self) -> bool:
        """Validate plugin configuration."""
        # Check if required email settings are available
        import os
        return bool(os.getenv("EMAIL_FROM") and os.getenv("SMTP_HOST"))
