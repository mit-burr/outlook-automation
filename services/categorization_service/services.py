# services/categorization_service/service.py
from enum import Enum
from typing import List, Dict, Set
import re
from services.outlook_service.models import Meeting

class MeetingCategory(str, Enum):
    COMPANY_WIDE = "Company-Wide"
    STAFF_TEAM = "Team/Staff"
    DEPARTMENT = "Department"
    ONBOARDING = "Onboarding"
    UNCATEGORIZED = "Uncategorized"

    @property
    def priority(self) -> int:
        """Return priority for tie-breaking (lower number = higher priority)"""
        priorities = {
            MeetingCategory.STAFF_TEAM: 1,
            MeetingCategory.DEPARTMENT: 2,
            MeetingCategory.COMPANY_WIDE: 3,
            MeetingCategory.ONBOARDING: 4,
            MeetingCategory.UNCATEGORIZED: 5
        }
        return priorities[self]

class CategorizationService:
    def __init__(self):
        self.category_keywords: Dict[MeetingCategory, Set[str]] = {
            MeetingCategory.COMPANY_WIDE: {
                # Direct company indicators
                "company", "corporate", "organization", "enterprise",
                # Large meetings
                "all hands", "all-hands", "town hall", "townhall",
                "summit", "quarterly", "annual", "yearly",
                # Company-wide events
                "announcement", "update", "briefing", "showcase",
                "celebration", "awards", "recognition",
                # Executive terms
                "executive", "leadership", "ceo", "cfo", "cto",
                # Company culture
                "culture", "values", "mission", "vision",
                # Company performance
                "earnings", "results", "performance", "strategy"
            },
            
            MeetingCategory.STAFF_TEAM: {
                # Team indicators
                "team", "staff", "squad", "crew", "group",
                # Regular meetings
                "standup", "sync", "check-in", "touchbase", "touch-base",
                "daily", "weekly", "biweekly", "monthly",
                # Small group indicators
                "1:1", "one on one", "one-on-one", "<>", 
                "catchup", "catch-up", "chat",
                # Team activities
                "huddle", "scrum", "alignment", "coordination",
                "status", "update", "sync", "collaboration",
                # Team roles
                "lead", "manager", "supervisor", "coordinator",
                "peer"
            },
            
            MeetingCategory.DEPARTMENT: {
                # Department names
                "engineering", "software", "development", "devops",
                "sales", "marketing", "finance", "accounting",
                "hr", "human resources", "support", "customer service",
                "operations", "it", "infrastructure", "security",
                "product", "design", "research", "qa",
                # Department activities
                "retrospective", "retro", "planning", "review",
                "sprint", "backlog", "roadmap", "strategy",
                "architecture", "deployment", "release",
                # Department-wide terms
                "department", "division", "unit", "branch",
                "initiative", "project", "program", "workflow",
                # Review/Planning terms
                "debrief", "analysis", "assessment", "evaluation",
                "quarterly review", "milestone", "objectives",
                "goals", "metrics", "kpi"
            },
            
            MeetingCategory.ONBOARDING: {
                # Direct indicators
                "onboarding", "orientation", "introduction", "intro",
                "new hire", "new-hire", "first day", "first week",
                # Common onboarding activities
                "training", "welcome", "overview", "setup",
                "documentation", "paperwork", "benefits",
                "getting started", "kickoff", "kick-off",
                # Onboarding related terms
                "mentor", "buddy", "guide", "tour",
                "handbook", "manual", "policies", "procedures",
                "system access", "credentials", "setup",
                # HR related
                "hr meeting", "employee", "policies", "i9",
                "direct deposit", "benefits", "enrollment"
            }
        }
        
        # Compile regex patterns for each category
        self.category_patterns: Dict[MeetingCategory, List[re.Pattern]] = {
            category: [
                re.compile(rf'\b{keyword}\b', re.IGNORECASE)
                for keyword in keywords
            ]
            for category, keywords in self.category_keywords.items()
        }

    def categorize_meeting(self, meeting: Meeting) -> MeetingCategory:
        """
        Categorize a single meeting based on its subject and other properties.
        Returns the most appropriate category based on keyword matches and priority.
        """
        search_text = f"{meeting.subject} {meeting.organizer}"
        
        # Track matches for each category
        matches = {category: 0 for category in MeetingCategory}
        
        # Count matches for each category
        for category, patterns in self.category_patterns.items():
            for pattern in patterns:
                if pattern.search(search_text):
                    matches[category] += 1
        
        # Find categories with the most matches
        max_matches = max(matches.values())
        if max_matches == 0:
            return MeetingCategory.UNCATEGORIZED
            
        # Get all categories that have the maximum number of matches
        top_categories = [
            category for category, count in matches.items() 
            if count == max_matches
        ]
        
        # If there's a tie, use priority to break it
        if len(top_categories) > 1:
            return min(top_categories, key=lambda x: x.priority)
        
        return top_categories[0]

    def categorize_meetings(self, meetings: List[Meeting]) -> Dict[MeetingCategory, List[Meeting]]:
        """
        Categorize a list of meetings and return them grouped by category.
        """
        categorized: Dict[MeetingCategory, List[Meeting]] = {
            category: [] for category in MeetingCategory
        }
        
        for meeting in meetings:
            category = self.categorize_meeting(meeting)
            categorized[category].append(meeting)
        
        return categorized

    def get_category_summary(self, categorized_meetings: Dict[MeetingCategory, List[Meeting]]) -> Dict[str, float]:
        """
        Generate a summary of time spent in each category.
        Returns a dictionary mapping category names to total hours.
        """
        summary = {}
        for category, meetings in categorized_meetings.items():
            total_minutes = sum(meeting.rounded_duration for meeting in meetings)
            summary[category] = total_minutes / 60  # Convert to hours
        return summary