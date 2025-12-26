#!/usr/bin/env python3
"""
Excel Forum Scraper - Download and organize discussions from major Excel forums.

Supported forums:
- MrExcel.com
- ExcelForum.com
- Chandoo.org

Usage:
    python excel_forum_scraper.py --forum mrexcel --output ./data --pages 10
    python excel_forum_scraper.py --forum all --output ./data --pages 5
"""

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, asdict, field
from datetime import datetime
from typing import List, Optional, Dict, Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from ratelimit import limits, sleep_and_retry

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class Post:
    """Represents a single post/reply in a discussion."""
    author: str
    content: str
    date: str
    post_number: int
    is_solution: bool = False


@dataclass
class Discussion:
    """Represents a forum discussion/thread."""
    title: str
    url: str
    author: str
    date_created: str
    category: str
    tags: List[str] = field(default_factory=list)
    views: int = 0
    replies: int = 0
    posts: List[Post] = field(default_factory=list)
    forum_source: str = ""
    is_solved: bool = False
    last_activity: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        data = asdict(self)
        data['posts'] = [asdict(p) for p in self.posts]
        return data


class BaseForumScraper(ABC):
    """Abstract base class for forum scrapers."""

    def __init__(self, output_dir: str = "./data", delay: float = 1.0):
        self.output_dir = output_dir
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        })
        self.discussions: List[Discussion] = []

    @property
    @abstractmethod
    def forum_name(self) -> str:
        """Return the forum name."""
        pass

    @property
    @abstractmethod
    def base_url(self) -> str:
        """Return the base URL of the forum."""
        pass

    @abstractmethod
    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Return list of forum sections with name and URL."""
        pass

    @abstractmethod
    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a section page."""
        pass

    @abstractmethod
    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single thread and return Discussion object."""
        pass

    @sleep_and_retry
    @limits(calls=30, period=60)
    def fetch_page(self, url: str) -> Optional[BeautifulSoup]:
        """Fetch a page with rate limiting and error handling."""
        try:
            time.sleep(self.delay)
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            return BeautifulSoup(response.text, 'lxml')
        except requests.RequestException as e:
            logger.error(f"Failed to fetch {url}: {e}")
            return None

    def scrape(self, max_pages: int = 5, max_threads_per_section: int = 50) -> List[Discussion]:
        """Main scraping method."""
        logger.info(f"Starting scrape of {self.forum_name}")

        sections = self.get_forum_sections()
        logger.info(f"Found {len(sections)} sections to scrape")

        for section in tqdm(sections, desc=f"Scraping {self.forum_name} sections"):
            section_name = section.get('name', 'Unknown')
            section_url = section.get('url', '')

            if not section_url:
                continue

            logger.info(f"Scraping section: {section_name}")
            threads_scraped = 0

            for page in range(1, max_pages + 1):
                if threads_scraped >= max_threads_per_section:
                    break

                threads = self.get_thread_list(section_url, page)

                if not threads:
                    break

                for thread_info in threads:
                    if threads_scraped >= max_threads_per_section:
                        break

                    thread_url = thread_info.get('url', '')
                    if not thread_url:
                        continue

                    discussion = self.scrape_thread(thread_url, section_name)
                    if discussion:
                        self.discussions.append(discussion)
                        threads_scraped += 1

        logger.info(f"Scraped {len(self.discussions)} discussions from {self.forum_name}")
        return self.discussions

    def save_to_json(self, filename: Optional[str] = None) -> str:
        """Save discussions to JSON file."""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.forum_name.lower().replace(' ', '_')}_{timestamp}.json"

        filepath = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        data = {
            'forum': self.forum_name,
            'scraped_at': datetime.now().isoformat(),
            'total_discussions': len(self.discussions),
            'discussions': [d.to_dict() for d in self.discussions]
        }

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"Saved {len(self.discussions)} discussions to {filepath}")
        return filepath

    def save_to_csv(self, filename: Optional[str] = None) -> str:
        """Save discussions to CSV file (summary without full posts)."""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.forum_name.lower().replace(' ', '_')}_{timestamp}.csv"

        filepath = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Title', 'URL', 'Author', 'Date Created', 'Category',
                'Tags', 'Views', 'Replies', 'Is Solved', 'Forum Source'
            ])

            for d in self.discussions:
                writer.writerow([
                    d.title, d.url, d.author, d.date_created, d.category,
                    ';'.join(d.tags), d.views, d.replies, d.is_solved, d.forum_source
                ])

        logger.info(f"Saved {len(self.discussions)} discussions to {filepath}")
        return filepath


class MrExcelScraper(BaseForumScraper):
    """Scraper for MrExcel.com forum."""

    @property
    def forum_name(self) -> str:
        return "MrExcel"

    @property
    def base_url(self) -> str:
        return "https://www.mrexcel.com/board/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get MrExcel forum sections."""
        sections = [
            {'name': 'Excel Questions', 'url': urljoin(self.base_url, 'forums/excel-questions.teknedia/')},
            {'name': 'VBA Macros', 'url': urljoin(self.base_url, 'forums/excel-questions.teknedia/')},
            {'name': 'Power Query', 'url': urljoin(self.base_url, 'forums/power-query.teknedia/')},
            {'name': 'Power Pivot', 'url': urljoin(self.base_url, 'forums/power-pivot.teknedia/')},
            {'name': 'Charts and Graphs', 'url': urljoin(self.base_url, 'forums/charts-and-charting-in-excel.teknedia/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a MrExcel section page."""
        url = section_url if page == 1 else f"{section_url}page-{page}"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.structItem--thread')

        for elem in thread_elements:
            title_elem = elem.select_one('.structItem-title a')
            if title_elem:
                thread_url = urljoin(self.base_url, title_elem.get('href', ''))
                threads.append({
                    'title': title_elem.get_text(strip=True),
                    'url': thread_url
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single MrExcel thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            title_elem = soup.select_one('.p-title-value')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            author_elem = soup.select_one('.message--post .message-userDetails a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            date_elem = soup.select_one('.message--post time')
            date_created = date_elem.get('datetime', '') if date_elem else ""

            # Get reply count
            reply_elem = soup.select_one('.p-description li:contains("Replies")')
            replies = 0
            if reply_elem:
                reply_text = reply_elem.get_text()
                match = re.search(r'(\d+)', reply_text)
                if match:
                    replies = int(match.group(1))

            # Get view count
            view_elem = soup.select_one('.p-description li:contains("Views")')
            views = 0
            if view_elem:
                view_text = view_elem.get_text()
                match = re.search(r'([\d,]+)', view_text)
                if match:
                    views = int(match.group(1).replace(',', ''))

            # Get tags
            tag_elems = soup.select('.tagItem')
            tags = [t.get_text(strip=True) for t in tag_elems]

            # Check if solved
            is_solved = bool(soup.select_one('.label--solved, .label--answered'))

            # Get posts
            posts = []
            post_elems = soup.select('.message--post')

            for idx, post_elem in enumerate(post_elems[:20], 1):  # Limit to first 20 posts
                post_author = post_elem.select_one('.message-userDetails a')
                post_content = post_elem.select_one('.message-body')
                post_date = post_elem.select_one('time')

                posts.append(Post(
                    author=post_author.get_text(strip=True) if post_author else "Unknown",
                    content=post_content.get_text(strip=True)[:5000] if post_content else "",
                    date=post_date.get('datetime', '') if post_date else "",
                    post_number=idx,
                    is_solution=bool(post_elem.select_one('.label--solution'))
                ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                date_created=date_created,
                category=category,
                tags=tags,
                views=views,
                replies=replies,
                posts=posts,
                forum_source=self.forum_name,
                is_solved=is_solved
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None


class ExcelForumScraper(BaseForumScraper):
    """Scraper for ExcelForum.com."""

    @property
    def forum_name(self) -> str:
        return "ExcelForum"

    @property
    def base_url(self) -> str:
        return "https://www.excelforum.com/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get ExcelForum sections."""
        sections = [
            {'name': 'Excel General', 'url': urljoin(self.base_url, 'excel-general/')},
            {'name': 'Excel Formulas', 'url': urljoin(self.base_url, 'excel-formulas-functions/')},
            {'name': 'Excel VBA', 'url': urljoin(self.base_url, 'excel-programming-vba-macros/')},
            {'name': 'Excel Charts', 'url': urljoin(self.base_url, 'excel-charting-pivots/')},
            {'name': 'Excel New Users', 'url': urljoin(self.base_url, 'excel-new-users-basics/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from an ExcelForum section page."""
        url = section_url if page == 1 else f"{section_url}page{page}/"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.threadtitle a, .topic-title a, .structItem-title a')

        for elem in thread_elements:
            href = elem.get('href', '')
            if href and '/threads/' in href or '/topic/' in href:
                thread_url = urljoin(self.base_url, href)
                threads.append({
                    'title': elem.get_text(strip=True),
                    'url': thread_url
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single ExcelForum thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            # Get title
            title_elem = soup.select_one('h1, .thread-title, .topic-title, .p-title-value')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            # Get author
            author_elem = soup.select_one('.postauthor a, .author a, .message-name a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            # Get date
            date_elem = soup.select_one('.postdate, .post-date, time')
            date_created = ""
            if date_elem:
                date_created = date_elem.get('datetime', '') or date_elem.get_text(strip=True)

            # Get stats from thread info
            stats_text = soup.get_text()

            views = 0
            view_match = re.search(r'Views?:?\s*([\d,]+)', stats_text, re.I)
            if view_match:
                views = int(view_match.group(1).replace(',', ''))

            replies = 0
            reply_match = re.search(r'Replies?:?\s*(\d+)', stats_text, re.I)
            if reply_match:
                replies = int(reply_match.group(1))

            # Get posts
            posts = []
            post_elems = soup.select('.post, .message, .postcontainer')[:20]

            for idx, post_elem in enumerate(post_elems, 1):
                post_author = post_elem.select_one('.username, .postauthor a, .message-name a')
                post_content = post_elem.select_one('.postcontent, .post-body, .message-body')
                post_date = post_elem.select_one('.postdate, time')

                if post_content:
                    posts.append(Post(
                        author=post_author.get_text(strip=True) if post_author else "Unknown",
                        content=post_content.get_text(strip=True)[:5000],
                        date=post_date.get_text(strip=True) if post_date else "",
                        post_number=idx
                    ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                date_created=date_created,
                category=category,
                views=views,
                replies=replies,
                posts=posts,
                forum_source=self.forum_name
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None


class ChandooScraper(BaseForumScraper):
    """Scraper for Chandoo.org forum."""

    @property
    def forum_name(self) -> str:
        return "Chandoo"

    @property
    def base_url(self) -> str:
        return "https://forum.chandoo.org/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get Chandoo forum sections."""
        sections = [
            {'name': 'Excel Help', 'url': urljoin(self.base_url, 'forums/excel-help.5/')},
            {'name': 'VBA Macros', 'url': urljoin(self.base_url, 'forums/vba-macros.6/')},
            {'name': 'Power Query', 'url': urljoin(self.base_url, 'forums/power-query.32/')},
            {'name': 'Charts', 'url': urljoin(self.base_url, 'forums/excel-charts.7/')},
            {'name': 'Dashboards', 'url': urljoin(self.base_url, 'forums/excel-dashboards.8/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a Chandoo section page."""
        url = section_url if page == 1 else f"{section_url}page-{page}"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.structItem--thread .structItem-title a')

        for elem in thread_elements:
            href = elem.get('href', '')
            if href:
                thread_url = urljoin(self.base_url, href)
                threads.append({
                    'title': elem.get_text(strip=True),
                    'url': thread_url
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single Chandoo thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            # Get title
            title_elem = soup.select_one('.p-title-value, h1')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            # Get author from first post
            author_elem = soup.select_one('.message--post .message-userDetails a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            # Get date
            date_elem = soup.select_one('.message--post time')
            date_created = date_elem.get('datetime', '') if date_elem else ""

            # Get reply count from page
            replies = len(soup.select('.message--post')) - 1

            # Get tags
            tag_elems = soup.select('.tagItem')
            tags = [t.get_text(strip=True) for t in tag_elems]

            # Check if solved
            is_solved = bool(soup.select_one('[class*="solved"], [class*="answered"]'))

            # Get posts
            posts = []
            post_elems = soup.select('.message--post')[:20]

            for idx, post_elem in enumerate(post_elems, 1):
                post_author = post_elem.select_one('.message-userDetails a')
                post_content = post_elem.select_one('.message-body .bbWrapper')
                post_date = post_elem.select_one('time')

                if post_content:
                    posts.append(Post(
                        author=post_author.get_text(strip=True) if post_author else "Unknown",
                        content=post_content.get_text(strip=True)[:5000],
                        date=post_date.get('datetime', '') if post_date else "",
                        post_number=idx,
                        is_solution=bool(post_elem.select_one('[class*="solution"]'))
                    ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                date_created=date_created,
                category=category,
                tags=tags,
                replies=max(0, replies),
                posts=posts,
                forum_source=self.forum_name,
                is_solved=is_solved
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None


class ForumOrganizer:
    """Organizes and analyzes scraped forum discussions."""

    def __init__(self, discussions: List[Discussion]):
        self.discussions = discussions

    def organize_by_category(self) -> Dict[str, List[Discussion]]:
        """Group discussions by category."""
        organized = {}
        for d in self.discussions:
            category = d.category or 'Uncategorized'
            if category not in organized:
                organized[category] = []
            organized[category].append(d)
        return organized

    def organize_by_forum(self) -> Dict[str, List[Discussion]]:
        """Group discussions by source forum."""
        organized = {}
        for d in self.discussions:
            forum = d.forum_source or 'Unknown'
            if forum not in organized:
                organized[forum] = []
            organized[forum].append(d)
        return organized

    def get_popular_topics(self, top_n: int = 20) -> List[Discussion]:
        """Get most popular discussions by view count."""
        return sorted(self.discussions, key=lambda x: x.views, reverse=True)[:top_n]

    def get_solved_discussions(self) -> List[Discussion]:
        """Get all solved/answered discussions."""
        return [d for d in self.discussions if d.is_solved]

    def get_statistics(self) -> Dict[str, Any]:
        """Get statistics about the scraped data."""
        total = len(self.discussions)
        solved = len([d for d in self.discussions if d.is_solved])

        total_views = sum(d.views for d in self.discussions)
        total_replies = sum(d.replies for d in self.discussions)

        categories = {}
        for d in self.discussions:
            cat = d.category or 'Uncategorized'
            categories[cat] = categories.get(cat, 0) + 1

        forums = {}
        for d in self.discussions:
            forum = d.forum_source or 'Unknown'
            forums[forum] = forums.get(forum, 0) + 1

        return {
            'total_discussions': total,
            'solved_discussions': solved,
            'solved_percentage': round(solved / total * 100, 2) if total else 0,
            'total_views': total_views,
            'total_replies': total_replies,
            'avg_views': round(total_views / total, 2) if total else 0,
            'avg_replies': round(total_replies / total, 2) if total else 0,
            'by_category': categories,
            'by_forum': forums
        }

    def save_organized_data(self, output_dir: str) -> None:
        """Save organized data to multiple files."""
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Save by category
        by_category = self.organize_by_category()
        for category, discussions in by_category.items():
            safe_name = re.sub(r'[^\w\-]', '_', category)
            filepath = os.path.join(output_dir, f"category_{safe_name}_{timestamp}.json")

            data = {
                'category': category,
                'count': len(discussions),
                'discussions': [d.to_dict() for d in discussions]
            }

            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

        # Save statistics
        stats = self.get_statistics()
        stats_path = os.path.join(output_dir, f"statistics_{timestamp}.json")
        with open(stats_path, 'w', encoding='utf-8') as f:
            json.dump(stats, f, indent=2)

        logger.info(f"Saved organized data to {output_dir}")


def main():
    """Main entry point for the scraper."""
    parser = argparse.ArgumentParser(
        description='Scrape and organize discussions from Excel forums',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --forum mrexcel --pages 5
  %(prog)s --forum all --output ./excel_data --pages 10
  %(prog)s --forum excelforum --threads 100
        """
    )

    parser.add_argument(
        '--forum',
        choices=['mrexcel', 'excelforum', 'chandoo', 'all'],
        default='all',
        help='Forum to scrape (default: all)'
    )
    parser.add_argument(
        '--output', '-o',
        default='./data',
        help='Output directory (default: ./data)'
    )
    parser.add_argument(
        '--pages', '-p',
        type=int,
        default=5,
        help='Max pages to scrape per section (default: 5)'
    )
    parser.add_argument(
        '--threads', '-t',
        type=int,
        default=50,
        help='Max threads per section (default: 50)'
    )
    parser.add_argument(
        '--delay', '-d',
        type=float,
        default=1.0,
        help='Delay between requests in seconds (default: 1.0)'
    )
    parser.add_argument(
        '--format', '-f',
        choices=['json', 'csv', 'both'],
        default='both',
        help='Output format (default: both)'
    )
    parser.add_argument(
        '--organize',
        action='store_true',
        help='Organize output by category'
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose logging'
    )

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Select scrapers
    scrapers = []
    if args.forum == 'all':
        scrapers = [
            MrExcelScraper(args.output, args.delay),
            ExcelForumScraper(args.output, args.delay),
            ChandooScraper(args.output, args.delay)
        ]
    elif args.forum == 'mrexcel':
        scrapers = [MrExcelScraper(args.output, args.delay)]
    elif args.forum == 'excelforum':
        scrapers = [ExcelForumScraper(args.output, args.delay)]
    elif args.forum == 'chandoo':
        scrapers = [ChandooScraper(args.output, args.delay)]

    all_discussions = []

    # Run scrapers
    for scraper in scrapers:
        try:
            discussions = scraper.scrape(
                max_pages=args.pages,
                max_threads_per_section=args.threads
            )
            all_discussions.extend(discussions)

            # Save individual forum data
            if args.format in ('json', 'both'):
                scraper.save_to_json()
            if args.format in ('csv', 'both'):
                scraper.save_to_csv()

        except KeyboardInterrupt:
            logger.info("Scraping interrupted by user")
            break
        except Exception as e:
            logger.error(f"Error scraping {scraper.forum_name}: {e}")

    # Organize if requested
    if args.organize and all_discussions:
        organizer = ForumOrganizer(all_discussions)
        organizer.save_organized_data(args.output)

        # Print statistics
        stats = organizer.get_statistics()
        print("\n=== Scraping Statistics ===")
        print(f"Total discussions: {stats['total_discussions']}")
        print(f"Solved discussions: {stats['solved_discussions']} ({stats['solved_percentage']}%)")
        print(f"Total views: {stats['total_views']:,}")
        print(f"Average views: {stats['avg_views']:,}")
        print(f"Average replies: {stats['avg_replies']}")
        print("\nBy Forum:")
        for forum, count in stats['by_forum'].items():
            print(f"  {forum}: {count}")
        print("\nBy Category:")
        for cat, count in sorted(stats['by_category'].items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {cat}: {count}")

    print(f"\nScraped {len(all_discussions)} total discussions")
    print(f"Data saved to: {args.output}")

    return 0


if __name__ == '__main__':
    sys.exit(main())
