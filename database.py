"""
Database module for Excel Forum Scraper.

Provides SQLite-based persistent storage for:
- Discussions/threads with update tracking
- User profiles
- Posts with history
- Scraping metadata
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple
from contextlib import contextmanager
from dataclasses import dataclass, asdict, field


@dataclass
class UserProfile:
    """Represents a forum user profile."""
    username: str
    user_id: str
    forum_source: str
    profile_url: str = ""
    join_date: str = ""
    post_count: int = 0
    reputation: int = 0
    location: str = ""
    title: str = ""
    avatar_url: str = ""
    signature: str = ""
    last_seen: str = ""
    bio: str = ""
    website: str = ""
    extra_data: Dict[str, Any] = field(default_factory=dict)
    first_scraped: str = ""
    last_updated: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        if isinstance(data.get('extra_data'), dict):
            data['extra_data'] = json.dumps(data['extra_data'])
        return data


class ForumDatabase:
    """SQLite database handler for forum data."""

    def __init__(self, db_path: str = "./data/forum_data.db"):
        self.db_path = db_path
        os.makedirs(os.path.dirname(db_path) if os.path.dirname(db_path) else ".", exist_ok=True)
        self._init_database()

    @contextmanager
    def get_connection(self):
        """Context manager for database connections."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_database(self):
        """Initialize database schema."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Discussions table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS discussions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    url TEXT UNIQUE NOT NULL,
                    title TEXT NOT NULL,
                    author TEXT,
                    author_id TEXT,
                    date_created TEXT,
                    category TEXT,
                    tags TEXT,
                    views INTEGER DEFAULT 0,
                    replies INTEGER DEFAULT 0,
                    forum_source TEXT NOT NULL,
                    is_solved INTEGER DEFAULT 0,
                    last_activity TEXT,
                    content_hash TEXT,
                    first_scraped TEXT NOT NULL,
                    last_updated TEXT NOT NULL,
                    update_count INTEGER DEFAULT 1,
                    is_active INTEGER DEFAULT 1
                )
            """)

            # Posts table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS posts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    discussion_id INTEGER NOT NULL,
                    post_number INTEGER NOT NULL,
                    author TEXT,
                    author_id TEXT,
                    content TEXT,
                    date TEXT,
                    is_solution INTEGER DEFAULT 0,
                    content_hash TEXT,
                    first_scraped TEXT NOT NULL,
                    last_updated TEXT NOT NULL,
                    FOREIGN KEY (discussion_id) REFERENCES discussions(id),
                    UNIQUE (discussion_id, post_number)
                )
            """)

            # User profiles table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL,
                    user_id TEXT,
                    forum_source TEXT NOT NULL,
                    profile_url TEXT,
                    join_date TEXT,
                    post_count INTEGER DEFAULT 0,
                    reputation INTEGER DEFAULT 0,
                    location TEXT,
                    title TEXT,
                    avatar_url TEXT,
                    signature TEXT,
                    last_seen TEXT,
                    bio TEXT,
                    website TEXT,
                    extra_data TEXT,
                    first_scraped TEXT NOT NULL,
                    last_updated TEXT NOT NULL,
                    UNIQUE (username, forum_source)
                )
            """)

            # Scrape history table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS scrape_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    forum_source TEXT NOT NULL,
                    section TEXT,
                    started_at TEXT NOT NULL,
                    completed_at TEXT,
                    threads_found INTEGER DEFAULT 0,
                    threads_new INTEGER DEFAULT 0,
                    threads_updated INTEGER DEFAULT 0,
                    users_found INTEGER DEFAULT 0,
                    status TEXT DEFAULT 'running',
                    error_message TEXT
                )
            """)

            # Thread tracking/watchlist table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS watched_threads (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    discussion_id INTEGER NOT NULL,
                    added_at TEXT NOT NULL,
                    last_checked TEXT,
                    check_frequency TEXT DEFAULT 'daily',
                    notify_on_update INTEGER DEFAULT 1,
                    notes TEXT,
                    FOREIGN KEY (discussion_id) REFERENCES discussions(id),
                    UNIQUE (discussion_id)
                )
            """)

            # Thread history for tracking changes over time
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS discussion_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    discussion_id INTEGER NOT NULL,
                    recorded_at TEXT NOT NULL,
                    views INTEGER,
                    replies INTEGER,
                    is_solved INTEGER,
                    FOREIGN KEY (discussion_id) REFERENCES discussions(id)
                )
            """)

            # Create indexes for common queries
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_discussions_url ON discussions(url)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_discussions_forum ON discussions(forum_source)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_discussions_category ON discussions(category)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_discussions_author ON discussions(author)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_discussions_updated ON discussions(last_updated)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_posts_discussion ON posts(discussion_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_forum ON users(forum_source)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")

    def _hash_content(self, content: str) -> str:
        """Generate a hash of content for change detection."""
        import hashlib
        return hashlib.md5(content.encode('utf-8')).hexdigest()

    # Discussion methods
    def get_discussion_by_url(self, url: str) -> Optional[Dict[str, Any]]:
        """Get a discussion by its URL."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM discussions WHERE url = ?", (url,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def upsert_discussion(self, discussion: Dict[str, Any]) -> Tuple[int, bool]:
        """
        Insert or update a discussion.
        Returns (discussion_id, is_new).
        """
        now = datetime.now().isoformat()
        url = discussion['url']

        # Generate content hash for change detection
        content_for_hash = f"{discussion.get('title', '')}{discussion.get('replies', 0)}{discussion.get('views', 0)}"
        content_hash = self._hash_content(content_for_hash)

        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Check if exists
            existing = self.get_discussion_by_url(url)

            if existing:
                # Check if content changed
                if existing.get('content_hash') == content_hash:
                    # No changes, just update last_checked
                    return existing['id'], False

                # Update existing
                cursor.execute("""
                    UPDATE discussions SET
                        title = ?,
                        views = ?,
                        replies = ?,
                        is_solved = ?,
                        last_activity = ?,
                        content_hash = ?,
                        last_updated = ?,
                        update_count = update_count + 1
                    WHERE url = ?
                """, (
                    discussion.get('title', existing['title']),
                    discussion.get('views', 0),
                    discussion.get('replies', 0),
                    1 if discussion.get('is_solved') else 0,
                    discussion.get('last_activity', ''),
                    content_hash,
                    now,
                    url
                ))

                # Record history
                cursor.execute("""
                    INSERT INTO discussion_history
                    (discussion_id, recorded_at, views, replies, is_solved)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    existing['id'],
                    now,
                    discussion.get('views', 0),
                    discussion.get('replies', 0),
                    1 if discussion.get('is_solved') else 0
                ))

                return existing['id'], False
            else:
                # Insert new
                tags = discussion.get('tags', [])
                if isinstance(tags, list):
                    tags = json.dumps(tags)

                cursor.execute("""
                    INSERT INTO discussions
                    (url, title, author, author_id, date_created, category, tags,
                     views, replies, forum_source, is_solved, last_activity,
                     content_hash, first_scraped, last_updated)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    url,
                    discussion.get('title', ''),
                    discussion.get('author', ''),
                    discussion.get('author_id', ''),
                    discussion.get('date_created', ''),
                    discussion.get('category', ''),
                    tags,
                    discussion.get('views', 0),
                    discussion.get('replies', 0),
                    discussion.get('forum_source', ''),
                    1 if discussion.get('is_solved') else 0,
                    discussion.get('last_activity', ''),
                    content_hash,
                    now,
                    now
                ))

                return cursor.lastrowid, True

    def upsert_post(self, discussion_id: int, post: Dict[str, Any]) -> Tuple[int, bool]:
        """Insert or update a post. Returns (post_id, is_new)."""
        now = datetime.now().isoformat()
        content_hash = self._hash_content(post.get('content', ''))

        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Check if exists
            cursor.execute("""
                SELECT id, content_hash FROM posts
                WHERE discussion_id = ? AND post_number = ?
            """, (discussion_id, post.get('post_number', 1)))
            existing = cursor.fetchone()

            if existing:
                if existing['content_hash'] == content_hash:
                    return existing['id'], False

                # Update
                cursor.execute("""
                    UPDATE posts SET
                        content = ?,
                        is_solution = ?,
                        content_hash = ?,
                        last_updated = ?
                    WHERE id = ?
                """, (
                    post.get('content', ''),
                    1 if post.get('is_solution') else 0,
                    content_hash,
                    now,
                    existing['id']
                ))
                return existing['id'], False
            else:
                # Insert
                cursor.execute("""
                    INSERT INTO posts
                    (discussion_id, post_number, author, author_id, content,
                     date, is_solution, content_hash, first_scraped, last_updated)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    discussion_id,
                    post.get('post_number', 1),
                    post.get('author', ''),
                    post.get('author_id', ''),
                    post.get('content', ''),
                    post.get('date', ''),
                    1 if post.get('is_solution') else 0,
                    content_hash,
                    now,
                    now
                ))
                return cursor.lastrowid, True

    # User methods
    def get_user(self, username: str, forum_source: str) -> Optional[Dict[str, Any]]:
        """Get a user by username and forum."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM users WHERE username = ? AND forum_source = ?",
                (username, forum_source)
            )
            row = cursor.fetchone()
            if row:
                data = dict(row)
                if data.get('extra_data'):
                    data['extra_data'] = json.loads(data['extra_data'])
                return data
            return None

    def upsert_user(self, user: Dict[str, Any]) -> Tuple[int, bool]:
        """Insert or update a user profile. Returns (user_id, is_new)."""
        now = datetime.now().isoformat()

        with self.get_connection() as conn:
            cursor = conn.cursor()

            existing = self.get_user(user['username'], user['forum_source'])

            extra_data = user.get('extra_data', {})
            if isinstance(extra_data, dict):
                extra_data = json.dumps(extra_data)

            if existing:
                cursor.execute("""
                    UPDATE users SET
                        profile_url = ?,
                        post_count = ?,
                        reputation = ?,
                        location = ?,
                        title = ?,
                        avatar_url = ?,
                        signature = ?,
                        last_seen = ?,
                        bio = ?,
                        website = ?,
                        extra_data = ?,
                        last_updated = ?
                    WHERE username = ? AND forum_source = ?
                """, (
                    user.get('profile_url', ''),
                    user.get('post_count', 0),
                    user.get('reputation', 0),
                    user.get('location', ''),
                    user.get('title', ''),
                    user.get('avatar_url', ''),
                    user.get('signature', ''),
                    user.get('last_seen', ''),
                    user.get('bio', ''),
                    user.get('website', ''),
                    extra_data,
                    now,
                    user['username'],
                    user['forum_source']
                ))
                return existing['id'], False
            else:
                cursor.execute("""
                    INSERT INTO users
                    (username, user_id, forum_source, profile_url, join_date,
                     post_count, reputation, location, title, avatar_url,
                     signature, last_seen, bio, website, extra_data,
                     first_scraped, last_updated)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    user['username'],
                    user.get('user_id', ''),
                    user['forum_source'],
                    user.get('profile_url', ''),
                    user.get('join_date', ''),
                    user.get('post_count', 0),
                    user.get('reputation', 0),
                    user.get('location', ''),
                    user.get('title', ''),
                    user.get('avatar_url', ''),
                    user.get('signature', ''),
                    user.get('last_seen', ''),
                    user.get('bio', ''),
                    user.get('website', ''),
                    extra_data,
                    now,
                    now
                ))
                return cursor.lastrowid, True

    # Watch/track threads
    def watch_thread(self, discussion_id: int, frequency: str = 'daily', notes: str = '') -> int:
        """Add a thread to watchlist."""
        now = datetime.now().isoformat()

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO watched_threads
                (discussion_id, added_at, check_frequency, notes)
                VALUES (?, ?, ?, ?)
            """, (discussion_id, now, frequency, notes))
            return cursor.lastrowid

    def unwatch_thread(self, discussion_id: int) -> bool:
        """Remove a thread from watchlist."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM watched_threads WHERE discussion_id = ?", (discussion_id,))
            return cursor.rowcount > 0

    def get_watched_threads(self) -> List[Dict[str, Any]]:
        """Get all watched threads with discussion details."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT w.*, d.url, d.title, d.views, d.replies, d.is_solved, d.last_updated
                FROM watched_threads w
                JOIN discussions d ON w.discussion_id = d.id
                ORDER BY w.added_at DESC
            """)
            return [dict(row) for row in cursor.fetchall()]

    def get_threads_needing_update(self, frequency: str = 'daily') -> List[Dict[str, Any]]:
        """Get watched threads that need to be checked based on frequency."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Calculate time threshold based on frequency
            if frequency == 'hourly':
                hours = 1
            elif frequency == 'daily':
                hours = 24
            elif frequency == 'weekly':
                hours = 168
            else:
                hours = 24

            cursor.execute("""
                SELECT w.*, d.url, d.title
                FROM watched_threads w
                JOIN discussions d ON w.discussion_id = d.id
                WHERE w.check_frequency = ?
                AND (w.last_checked IS NULL
                     OR datetime(w.last_checked) < datetime('now', ?))
            """, (frequency, f'-{hours} hours'))

            return [dict(row) for row in cursor.fetchall()]

    # Scrape history
    def start_scrape_session(self, forum_source: str, section: str = '') -> int:
        """Record the start of a scrape session."""
        now = datetime.now().isoformat()
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO scrape_history (forum_source, section, started_at, status)
                VALUES (?, ?, ?, 'running')
            """, (forum_source, section, now))
            return cursor.lastrowid

    def end_scrape_session(self, session_id: int, threads_found: int = 0,
                           threads_new: int = 0, threads_updated: int = 0,
                           users_found: int = 0, error: str = None):
        """Record the end of a scrape session."""
        now = datetime.now().isoformat()
        status = 'error' if error else 'completed'

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE scrape_history SET
                    completed_at = ?,
                    threads_found = ?,
                    threads_new = ?,
                    threads_updated = ?,
                    users_found = ?,
                    status = ?,
                    error_message = ?
                WHERE id = ?
            """, (now, threads_found, threads_new, threads_updated,
                  users_found, status, error, session_id))

    # Query methods
    def get_discussions(self, forum_source: str = None, category: str = None,
                        limit: int = 100, offset: int = 0,
                        order_by: str = 'last_updated') -> List[Dict[str, Any]]:
        """Get discussions with optional filters."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            query = "SELECT * FROM discussions WHERE 1=1"
            params = []

            if forum_source:
                query += " AND forum_source = ?"
                params.append(forum_source)

            if category:
                query += " AND category = ?"
                params.append(category)

            query += f" ORDER BY {order_by} DESC LIMIT ? OFFSET ?"
            params.extend([limit, offset])

            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_discussion_with_posts(self, discussion_id: int) -> Optional[Dict[str, Any]]:
        """Get a discussion with all its posts."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM discussions WHERE id = ?", (discussion_id,))
            discussion = cursor.fetchone()

            if not discussion:
                return None

            result = dict(discussion)

            cursor.execute(
                "SELECT * FROM posts WHERE discussion_id = ? ORDER BY post_number",
                (discussion_id,)
            )
            result['posts'] = [dict(row) for row in cursor.fetchall()]

            return result

    def get_discussion_history(self, discussion_id: int) -> List[Dict[str, Any]]:
        """Get historical data for a discussion."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM discussion_history
                WHERE discussion_id = ?
                ORDER BY recorded_at DESC
            """, (discussion_id,))
            return [dict(row) for row in cursor.fetchall()]

    def get_users(self, forum_source: str = None, limit: int = 100,
                  offset: int = 0) -> List[Dict[str, Any]]:
        """Get user profiles with optional forum filter."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            if forum_source:
                cursor.execute("""
                    SELECT * FROM users
                    WHERE forum_source = ?
                    ORDER BY post_count DESC
                    LIMIT ? OFFSET ?
                """, (forum_source, limit, offset))
            else:
                cursor.execute("""
                    SELECT * FROM users
                    ORDER BY post_count DESC
                    LIMIT ? OFFSET ?
                """, (limit, offset))

            return [dict(row) for row in cursor.fetchall()]

    def get_statistics(self) -> Dict[str, Any]:
        """Get overall database statistics."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            stats = {}

            cursor.execute("SELECT COUNT(*) FROM discussions")
            stats['total_discussions'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM discussions WHERE is_solved = 1")
            stats['solved_discussions'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM posts")
            stats['total_posts'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM users")
            stats['total_users'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM watched_threads")
            stats['watched_threads'] = cursor.fetchone()[0]

            cursor.execute("""
                SELECT forum_source, COUNT(*) as count
                FROM discussions
                GROUP BY forum_source
            """)
            stats['by_forum'] = {row['forum_source']: row['count']
                                 for row in cursor.fetchall()}

            cursor.execute("""
                SELECT category, COUNT(*) as count
                FROM discussions
                GROUP BY category
                ORDER BY count DESC
                LIMIT 10
            """)
            stats['top_categories'] = {row['category']: row['count']
                                       for row in cursor.fetchall()}

            cursor.execute("SELECT MAX(last_updated) FROM discussions")
            stats['last_update'] = cursor.fetchone()[0]

            return stats

    def search_discussions(self, query: str, limit: int = 50) -> List[Dict[str, Any]]:
        """Search discussions by title or content."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            search_term = f"%{query}%"

            cursor.execute("""
                SELECT DISTINCT d.* FROM discussions d
                LEFT JOIN posts p ON d.id = p.discussion_id
                WHERE d.title LIKE ? OR p.content LIKE ?
                ORDER BY d.views DESC
                LIMIT ?
            """, (search_term, search_term, limit))

            return [dict(row) for row in cursor.fetchall()]

    def get_new_discussions_since(self, since: str, forum_source: str = None) -> List[Dict[str, Any]]:
        """Get discussions added since a given timestamp."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            if forum_source:
                cursor.execute("""
                    SELECT * FROM discussions
                    WHERE first_scraped > ? AND forum_source = ?
                    ORDER BY first_scraped DESC
                """, (since, forum_source))
            else:
                cursor.execute("""
                    SELECT * FROM discussions
                    WHERE first_scraped > ?
                    ORDER BY first_scraped DESC
                """, (since,))

            return [dict(row) for row in cursor.fetchall()]

    def get_updated_discussions_since(self, since: str, forum_source: str = None) -> List[Dict[str, Any]]:
        """Get discussions updated since a given timestamp."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            if forum_source:
                cursor.execute("""
                    SELECT * FROM discussions
                    WHERE last_updated > ? AND first_scraped < ? AND forum_source = ?
                    ORDER BY last_updated DESC
                """, (since, since, forum_source))
            else:
                cursor.execute("""
                    SELECT * FROM discussions
                    WHERE last_updated > ? AND first_scraped < ?
                    ORDER BY last_updated DESC
                """, (since, since))

            return [dict(row) for row in cursor.fetchall()]

    def export_to_json(self, filepath: str, forum_source: str = None):
        """Export all data to a JSON file."""
        data = {
            'exported_at': datetime.now().isoformat(),
            'statistics': self.get_statistics(),
            'discussions': [],
            'users': self.get_users(forum_source, limit=10000)
        }

        discussions = self.get_discussions(forum_source, limit=10000)
        for d in discussions:
            full_discussion = self.get_discussion_with_posts(d['id'])
            if full_discussion:
                data['discussions'].append(full_discussion)

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        return len(data['discussions'])
