# memory_system.py — Robust conversation memory system
# - Fast in-memory caching with LRU eviction
# - Persistent SQLite storage for long-term memory
# - Phone number and email conversation tracking
# - Efficient search and retrieval
# - Thread-safe operations

import sqlite3
import json
import threading
import time
from typing import Dict, List, Optional, Tuple, Any
from collections import OrderedDict
from datetime import datetime, timezone
import hashlib
import os

class ConversationMemory:
    """Robust conversation memory system with caching and persistence."""
    
    def __init__(self, db_path: str = "conversations.db", cache_size: int = 1000):
        self.db_path = db_path
        self.cache_size = cache_size
        self.lock = threading.RLock()
        
        # In-memory LRU cache: {user_id: [conversations]}
        self.cache: OrderedDict[str, List[Dict]] = OrderedDict()
        
        # Initialize database
        self._init_database()
        
        # Load recent conversations into cache
        self._load_recent_conversations()
    
    def _init_database(self):
        """Initialize SQLite database with conversation tables."""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS conversations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id TEXT NOT NULL,
                    user_type TEXT NOT NULL,  -- 'phone' or 'email'
                    message TEXT NOT NULL,
                    response TEXT NOT NULL,
                    timestamp REAL NOT NULL,
                    metadata TEXT,  -- JSON metadata
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_user_id ON conversations(user_id)
            """)
            
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_timestamp ON conversations(timestamp)
            """)
            
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_user_type ON conversations(user_type)
            """)
    
    def _load_recent_conversations(self):
        """Load recent conversations into cache for fast access."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("""
                SELECT user_id, user_type, message, response, timestamp, metadata
                FROM conversations 
                ORDER BY timestamp DESC 
                LIMIT ?
            """, (self.cache_size * 2,))
            
            for row in cursor.fetchall():
                user_id, user_type, message, response, timestamp, metadata = row
                conversation = {
                    'message': message,
                    'response': response,
                    'timestamp': timestamp,
                    'metadata': json.loads(metadata) if metadata else {}
                }
                
                if user_id not in self.cache:
                    self.cache[user_id] = []
                self.cache[user_id].append(conversation)
                
                # Move to end (most recent)
                self.cache.move_to_end(user_id)
    
    def _evict_oldest(self):
        """Evict oldest entries from cache when it gets too large."""
        while len(self.cache) > self.cache_size:
            # Remove oldest user's conversations
            user_id, conversations = self.cache.popitem(last=False)
            # Keep only the most recent conversation for this user
            if conversations:
                self.cache[user_id] = [conversations[-1]]
                self.cache.move_to_end(user_id)
    
    def store_conversation(self, user_id: str, user_type: str, message: str, 
                          response: str, metadata: Optional[Dict] = None) -> None:
        """Store a conversation in both cache and database."""
        with self.lock:
            timestamp = time.time()
            metadata = metadata or {}
            
            conversation = {
                'message': message,
                'response': response,
                'timestamp': timestamp,
                'metadata': metadata
            }
            
            # Add to cache
            if user_id not in self.cache:
                self.cache[user_id] = []
            self.cache[user_id].append(conversation)
            self.cache.move_to_end(user_id)
            
            # Evict if cache is too large
            if len(self.cache) > self.cache_size:
                self._evict_oldest()
            
            # Store in database
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    INSERT INTO conversations (user_id, user_type, message, response, timestamp, metadata)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (user_id, user_type, message, response, timestamp, json.dumps(metadata)))
    
    def get_conversations(self, user_id: str, limit: int = 10) -> List[Dict]:
        """Get recent conversations for a user."""
        with self.lock:
            # Try cache first
            if user_id in self.cache:
                conversations = self.cache[user_id][-limit:]
                self.cache.move_to_end(user_id)  # Mark as recently used
                return conversations
            
            # Fallback to database
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute("""
                    SELECT message, response, timestamp, metadata
                    FROM conversations 
                    WHERE user_id = ? 
                    ORDER BY timestamp DESC 
                    LIMIT ?
                """, (user_id, limit))
                
                conversations = []
                for row in cursor.fetchall():
                    message, response, timestamp, metadata = row
                    conversations.append({
                        'message': message,
                        'response': response,
                        'timestamp': timestamp,
                        'metadata': json.loads(metadata) if metadata else {}
                    })
                
                # Add to cache
                if conversations:
                    self.cache[user_id] = conversations
                    self.cache.move_to_end(user_id)
                
                return conversations
    
    def get_conversation_context(self, user_id: str, max_context: int = 5) -> str:
        """Get conversation context as a formatted string for AI."""
        conversations = self.get_conversations(user_id, max_context)
        
        if not conversations:
            return ""
        
        context_parts = []
        for conv in reversed(conversations):  # Oldest first
            timestamp = datetime.fromtimestamp(conv['timestamp']).strftime('%H:%M')
            context_parts.append(f"[{timestamp}] User: {conv['message']}")
            context_parts.append(f"[{timestamp}] AI: {conv['response']}")
        
        return "\n".join(context_parts)
    
    def search_conversations(self, query: str, user_id: Optional[str] = None, 
                           limit: int = 20) -> List[Dict]:
        """Search conversations by content."""
        with sqlite3.connect(self.db_path) as conn:
            if user_id:
                cursor = conn.execute("""
                    SELECT user_id, user_type, message, response, timestamp, metadata
                    FROM conversations 
                    WHERE user_id = ? AND (message LIKE ? OR response LIKE ?)
                    ORDER BY timestamp DESC 
                    LIMIT ?
                """, (user_id, f"%{query}%", f"%{query}%", limit))
            else:
                cursor = conn.execute("""
                    SELECT user_id, user_type, message, response, timestamp, metadata
                    FROM conversations 
                    WHERE message LIKE ? OR response LIKE ?
                    ORDER BY timestamp DESC 
                    LIMIT ?
                """, (f"%{query}%", f"%{query}%", limit))
            
            results = []
            for row in cursor.fetchall():
                user_id, user_type, message, response, timestamp, metadata = row
                results.append({
                    'user_id': user_id,
                    'user_type': user_type,
                    'message': message,
                    'response': response,
                    'timestamp': timestamp,
                    'metadata': json.loads(metadata) if metadata else {}
                })
            
            return results
    
    def get_user_stats(self, user_id: str) -> Dict[str, Any]:
        """Get statistics for a specific user."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("""
                SELECT 
                    COUNT(*) as total_conversations,
                    MIN(timestamp) as first_conversation,
                    MAX(timestamp) as last_conversation,
                    user_type
                FROM conversations 
                WHERE user_id = ?
                GROUP BY user_type
            """, (user_id,))
            
            stats = {
                'user_id': user_id,
                'total_conversations': 0,
                'first_conversation': None,
                'last_conversation': None,
                'by_type': {}
            }
            
            for row in cursor.fetchall():
                count, first, last, user_type = row
                stats['total_conversations'] += count
                if not stats['first_conversation'] or first < stats['first_conversation']:
                    stats['first_conversation'] = first
                if not stats['last_conversation'] or last > stats['last_conversation']:
                    stats['last_conversation'] = last
                stats['by_type'][user_type] = {
                    'count': count,
                    'first': first,
                    'last': last
                }
            
            return stats
    
    def get_system_stats(self) -> Dict[str, Any]:
        """Get overall system statistics."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("""
                SELECT 
                    COUNT(*) as total_conversations,
                    COUNT(DISTINCT user_id) as unique_users,
                    MIN(timestamp) as first_conversation,
                    MAX(timestamp) as last_conversation
                FROM conversations
            """)
            
            row = cursor.fetchone()
            if row:
                total, unique_users, first, last = row
                return {
                    'total_conversations': total,
                    'unique_users': unique_users,
                    'first_conversation': first,
                    'last_conversation': last,
                    'cache_size': len(self.cache),
                    'cache_users': len(self.cache)
                }
            return {}
    
    def clear_user_conversations(self, user_id: str) -> int:
        """Clear all conversations for a specific user."""
        with self.lock:
            # Remove from cache
            if user_id in self.cache:
                del self.cache[user_id]
            
            # Remove from database
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute("DELETE FROM conversations WHERE user_id = ?", (user_id,))
                return cursor.rowcount
    
    def clear_old_conversations(self, days_old: int = 30) -> int:
        """Clear conversations older than specified days."""
        cutoff_time = time.time() - (days_old * 24 * 60 * 60)
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("DELETE FROM conversations WHERE timestamp < ?", (cutoff_time,))
            return cursor.rowcount
    
    def export_conversations(self, user_id: Optional[str] = None, 
                           format: str = 'json') -> str:
        """Export conversations to JSON or CSV format."""
        if user_id:
            conversations = self.get_conversations(user_id, limit=1000)
        else:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute("""
                    SELECT user_id, user_type, message, response, timestamp, metadata
                    FROM conversations 
                    ORDER BY timestamp DESC
                """)
                conversations = []
                for row in cursor.fetchall():
                    user_id, user_type, message, response, timestamp, metadata = row
                    conversations.append({
                        'user_id': user_id,
                        'user_type': user_type,
                        'message': message,
                        'response': response,
                        'timestamp': timestamp,
                        'metadata': json.loads(metadata) if metadata else {}
                    })
        
        if format.lower() == 'json':
            return json.dumps(conversations, indent=2, default=str)
        elif format.lower() == 'csv':
            import csv
            import io
            output = io.StringIO()
            if conversations:
                writer = csv.DictWriter(output, fieldnames=conversations[0].keys())
                writer.writeheader()
                writer.writerows(conversations)
            return output.getvalue()
        else:
            raise ValueError("Format must be 'json' or 'csv'")

# Global memory instance
_memory = None

def get_memory() -> ConversationMemory:
    """Get the global memory instance."""
    global _memory
    if _memory is None:
        _memory = ConversationMemory()
    return _memory

def store_conversation(user_id: str, user_type: str, message: str, 
                      response: str, metadata: Optional[Dict] = None) -> None:
    """Store a conversation (convenience function)."""
    get_memory().store_conversation(user_id, user_type, message, response, metadata)

def get_conversation_context(user_id: str, max_context: int = 5) -> str:
    """Get conversation context (convenience function)."""
    return get_memory().get_conversation_context(user_id, max_context)

def search_conversations(query: str, user_id: Optional[str] = None, 
                        limit: int = 20) -> List[Dict]:
    """Search conversations (convenience function)."""
    return get_memory().search_conversations(query, user_id, limit)
