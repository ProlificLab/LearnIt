"""
Async Scraper Module for Excel Forum Scraper.

Features:
- Asynchronous scraping with aiohttp for high performance
- Proxy rotation with health checking
- Resumable sessions with checkpoints
- Concurrent request limiting
- Automatic retry with exponential backoff
"""

import asyncio
import aiohttp
import json
import logging
import os
import random
import time
import hashlib
from abc import ABC, abstractmethod
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Set, Tuple, Callable
from urllib.parse import urljoin, urlparse
from contextlib import asynccontextmanager

from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)


@dataclass
class ProxyConfig:
    """Configuration for a proxy."""
    url: str
    protocol: str = "http"
    username: Optional[str] = None
    password: Optional[str] = None
    last_used: Optional[datetime] = None
    success_count: int = 0
    failure_count: int = 0
    avg_response_time: float = 0.0
    is_healthy: bool = True

    @property
    def auth(self) -> Optional[aiohttp.BasicAuth]:
        if self.username and self.password:
            return aiohttp.BasicAuth(self.username, self.password)
        return None

    @property
    def full_url(self) -> str:
        if self.username and self.password:
            parsed = urlparse(self.url)
            return f"{parsed.scheme}://{self.username}:{self.password}@{parsed.netloc}"
        return self.url

    @property
    def success_rate(self) -> float:
        total = self.success_count + self.failure_count
        return self.success_count / total if total > 0 else 0.5


@dataclass
class ScrapeCheckpoint:
    """Checkpoint for resumable scraping."""
    session_id: str
    forum: str
    started_at: str
    last_updated: str
    current_section: str
    current_page: int
    completed_sections: List[str] = field(default_factory=list)
    completed_urls: Set[str] = field(default_factory=set)
    failed_urls: Dict[str, int] = field(default_factory=dict)  # url -> retry count
    stats: Dict[str, int] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'session_id': self.session_id,
            'forum': self.forum,
            'started_at': self.started_at,
            'last_updated': self.last_updated,
            'current_section': self.current_section,
            'current_page': self.current_page,
            'completed_sections': self.completed_sections,
            'completed_urls': list(self.completed_urls),
            'failed_urls': self.failed_urls,
            'stats': self.stats
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ScrapeCheckpoint':
        return cls(
            session_id=data['session_id'],
            forum=data['forum'],
            started_at=data['started_at'],
            last_updated=data['last_updated'],
            current_section=data['current_section'],
            current_page=data['current_page'],
            completed_sections=data.get('completed_sections', []),
            completed_urls=set(data.get('completed_urls', [])),
            failed_urls=data.get('failed_urls', {}),
            stats=data.get('stats', {})
        )

    def save(self, filepath: str):
        """Save checkpoint to file."""
        self.last_updated = datetime.now().isoformat()
        with open(filepath, 'w') as f:
            json.dump(self.to_dict(), f, indent=2)

    @classmethod
    def load(cls, filepath: str) -> Optional['ScrapeCheckpoint']:
        """Load checkpoint from file."""
        if not os.path.exists(filepath):
            return None
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except (json.JSONDecodeError, KeyError) as e:
            logger.error(f"Failed to load checkpoint: {e}")
            return None


class ProxyRotator:
    """Manage and rotate proxies with health checking."""

    def __init__(self, proxies: List[str] = None, check_url: str = "https://httpbin.org/ip"):
        self.proxies: List[ProxyConfig] = []
        self.check_url = check_url
        self.min_success_rate = 0.3
        self.cooldown_period = timedelta(minutes=5)
        self._lock = asyncio.Lock()

        if proxies:
            for proxy in proxies:
                self.add_proxy(proxy)

    def add_proxy(self, proxy_url: str, username: str = None, password: str = None):
        """Add a proxy to the rotation pool."""
        config = ProxyConfig(
            url=proxy_url,
            username=username,
            password=password
        )
        self.proxies.append(config)

    def load_from_file(self, filepath: str):
        """Load proxies from a file (one per line, format: url or url:user:pass)."""
        if not os.path.exists(filepath):
            return

        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue

                parts = line.split(':')
                if len(parts) >= 4:  # protocol://host:port:user:pass
                    url = ':'.join(parts[:3])
                    username = parts[3]
                    password = ':'.join(parts[4:]) if len(parts) > 4 else None
                    self.add_proxy(url, username, password)
                else:
                    self.add_proxy(line)

    async def get_proxy(self) -> Optional[ProxyConfig]:
        """Get the next available proxy."""
        async with self._lock:
            if not self.proxies:
                return None

            # Filter healthy proxies
            now = datetime.now()
            available = [
                p for p in self.proxies
                if p.is_healthy and (
                    p.last_used is None or
                    now - p.last_used > timedelta(seconds=1)
                )
            ]

            if not available:
                # Try proxies that were marked unhealthy but cooldown expired
                available = [
                    p for p in self.proxies
                    if not p.is_healthy and p.last_used and
                    now - p.last_used > self.cooldown_period
                ]
                for p in available:
                    p.is_healthy = True  # Give them another chance

            if not available:
                # Return the least recently used proxy
                available = sorted(self.proxies, key=lambda p: p.last_used or datetime.min)

            if not available:
                return None

            # Weight by success rate
            weights = [max(0.1, p.success_rate) for p in available]
            total_weight = sum(weights)
            weights = [w / total_weight for w in weights]

            proxy = random.choices(available, weights=weights, k=1)[0]
            proxy.last_used = now
            return proxy

    async def report_success(self, proxy: ProxyConfig, response_time: float):
        """Report successful request."""
        async with self._lock:
            proxy.success_count += 1
            proxy.is_healthy = True

            # Update average response time
            total = proxy.success_count + proxy.failure_count
            proxy.avg_response_time = (
                (proxy.avg_response_time * (total - 1) + response_time) / total
            )

    async def report_failure(self, proxy: ProxyConfig):
        """Report failed request."""
        async with self._lock:
            proxy.failure_count += 1

            # Mark as unhealthy if success rate drops too low
            if proxy.success_rate < self.min_success_rate and proxy.failure_count >= 3:
                proxy.is_healthy = False
                logger.warning(f"Proxy {proxy.url} marked as unhealthy")

    async def health_check(self, session: aiohttp.ClientSession) -> Dict[str, bool]:
        """Check health of all proxies."""
        results = {}

        async def check_one(proxy: ProxyConfig):
            try:
                start = time.time()
                async with session.get(
                    self.check_url,
                    proxy=proxy.full_url,
                    timeout=aiohttp.ClientTimeout(total=10)
                ) as response:
                    if response.status == 200:
                        elapsed = time.time() - start
                        await self.report_success(proxy, elapsed)
                        results[proxy.url] = True
                    else:
                        await self.report_failure(proxy)
                        results[proxy.url] = False
            except Exception:
                await self.report_failure(proxy)
                results[proxy.url] = False

        await asyncio.gather(*[check_one(p) for p in self.proxies])
        return results

    def get_stats(self) -> List[Dict[str, Any]]:
        """Get statistics for all proxies."""
        return [
            {
                'url': p.url,
                'is_healthy': p.is_healthy,
                'success_count': p.success_count,
                'failure_count': p.failure_count,
                'success_rate': round(p.success_rate, 2),
                'avg_response_time': round(p.avg_response_time, 2)
            }
            for p in self.proxies
        ]


class AsyncRateLimiter:
    """Async rate limiter with token bucket algorithm."""

    def __init__(self, rate: float = 1.0, burst: int = 5):
        """
        Args:
            rate: Requests per second
            burst: Maximum burst size
        """
        self.rate = rate
        self.burst = burst
        self.tokens = burst
        self.last_update = time.monotonic()
        self._lock = asyncio.Lock()

    async def acquire(self):
        """Acquire a token, waiting if necessary."""
        async with self._lock:
            now = time.monotonic()
            elapsed = now - self.last_update
            self.last_update = now

            # Add tokens based on elapsed time
            self.tokens = min(self.burst, self.tokens + elapsed * self.rate)

            if self.tokens < 1:
                # Wait for a token
                wait_time = (1 - self.tokens) / self.rate
                await asyncio.sleep(wait_time)
                self.tokens = 0
            else:
                self.tokens -= 1


class AsyncScraper:
    """High-performance async scraper with proxy rotation and checkpoints."""

    def __init__(
        self,
        max_concurrent: int = 10,
        requests_per_second: float = 2.0,
        timeout: int = 30,
        max_retries: int = 3,
        checkpoint_dir: str = "./checkpoints",
        proxy_file: str = None
    ):
        self.max_concurrent = max_concurrent
        self.timeout = aiohttp.ClientTimeout(total=timeout)
        self.max_retries = max_retries
        self.checkpoint_dir = checkpoint_dir
        self.rate_limiter = AsyncRateLimiter(rate=requests_per_second, burst=5)
        self.proxy_rotator = ProxyRotator()

        if proxy_file:
            self.proxy_rotator.load_from_file(proxy_file)

        self.semaphore = asyncio.Semaphore(max_concurrent)
        self.session: Optional[aiohttp.ClientSession] = None

        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
        }

        os.makedirs(checkpoint_dir, exist_ok=True)

    @asynccontextmanager
    async def get_session(self):
        """Context manager for aiohttp session."""
        if self.session is None or self.session.closed:
            connector = aiohttp.TCPConnector(
                limit=self.max_concurrent,
                limit_per_host=5,
                ttl_dns_cache=300
            )
            self.session = aiohttp.ClientSession(
                connector=connector,
                headers=self.headers,
                timeout=self.timeout
            )

        try:
            yield self.session
        finally:
            pass  # Keep session open for reuse

    async def close(self):
        """Close the session."""
        if self.session and not self.session.closed:
            await self.session.close()

    async def fetch(self, url: str, use_proxy: bool = True) -> Optional[str]:
        """Fetch a URL with rate limiting and optional proxy."""
        await self.rate_limiter.acquire()

        proxy = None
        proxy_config = None
        if use_proxy and self.proxy_rotator.proxies:
            proxy_config = await self.proxy_rotator.get_proxy()
            if proxy_config:
                proxy = proxy_config.full_url

        async with self.semaphore:
            for attempt in range(self.max_retries):
                try:
                    start_time = time.time()

                    async with self.get_session() as session:
                        async with session.get(url, proxy=proxy) as response:
                            if response.status == 200:
                                html = await response.text()

                                if proxy_config:
                                    elapsed = time.time() - start_time
                                    await self.proxy_rotator.report_success(proxy_config, elapsed)

                                return html

                            elif response.status == 429:  # Rate limited
                                wait_time = int(response.headers.get('Retry-After', 60))
                                logger.warning(f"Rate limited, waiting {wait_time}s")
                                await asyncio.sleep(wait_time)

                            elif response.status >= 500:
                                # Server error, retry
                                if proxy_config:
                                    await self.proxy_rotator.report_failure(proxy_config)

                            else:
                                logger.error(f"HTTP {response.status} for {url}")
                                return None

                except asyncio.TimeoutError:
                    logger.warning(f"Timeout fetching {url} (attempt {attempt + 1})")
                    if proxy_config:
                        await self.proxy_rotator.report_failure(proxy_config)

                except aiohttp.ClientError as e:
                    logger.warning(f"Client error fetching {url}: {e} (attempt {attempt + 1})")
                    if proxy_config:
                        await self.proxy_rotator.report_failure(proxy_config)

                except Exception as e:
                    logger.error(f"Error fetching {url}: {e}")
                    if proxy_config:
                        await self.proxy_rotator.report_failure(proxy_config)

                # Exponential backoff
                if attempt < self.max_retries - 1:
                    wait_time = (2 ** attempt) + random.uniform(0, 1)
                    await asyncio.sleep(wait_time)

                    # Try a different proxy on retry
                    if use_proxy and self.proxy_rotator.proxies:
                        proxy_config = await self.proxy_rotator.get_proxy()
                        proxy = proxy_config.full_url if proxy_config else None

        return None

    async def fetch_many(self, urls: List[str], use_proxy: bool = True,
                         callback: Callable[[str, str], None] = None) -> Dict[str, Optional[str]]:
        """Fetch multiple URLs concurrently."""
        results = {}

        async def fetch_one(url: str):
            html = await self.fetch(url, use_proxy)
            results[url] = html
            if callback and html:
                callback(url, html)

        await asyncio.gather(*[fetch_one(url) for url in urls])
        return results

    def get_checkpoint_path(self, forum: str) -> str:
        """Get checkpoint file path for a forum."""
        return os.path.join(self.checkpoint_dir, f"{forum.lower()}_checkpoint.json")

    def create_checkpoint(self, forum: str) -> ScrapeCheckpoint:
        """Create a new checkpoint."""
        return ScrapeCheckpoint(
            session_id=hashlib.md5(f"{forum}{datetime.now().isoformat()}".encode()).hexdigest()[:12],
            forum=forum,
            started_at=datetime.now().isoformat(),
            last_updated=datetime.now().isoformat(),
            current_section="",
            current_page=1,
            stats={'threads_found': 0, 'threads_new': 0, 'threads_updated': 0, 'errors': 0}
        )

    def load_or_create_checkpoint(self, forum: str) -> ScrapeCheckpoint:
        """Load existing checkpoint or create new one."""
        checkpoint_path = self.get_checkpoint_path(forum)
        checkpoint = ScrapeCheckpoint.load(checkpoint_path)

        if checkpoint:
            logger.info(f"Resuming from checkpoint: {checkpoint.session_id}")
            logger.info(f"  Completed sections: {len(checkpoint.completed_sections)}")
            logger.info(f"  Completed URLs: {len(checkpoint.completed_urls)}")
        else:
            checkpoint = self.create_checkpoint(forum)
            logger.info(f"Created new checkpoint: {checkpoint.session_id}")

        return checkpoint

    def save_checkpoint(self, checkpoint: ScrapeCheckpoint):
        """Save checkpoint to file."""
        checkpoint_path = self.get_checkpoint_path(checkpoint.forum)
        checkpoint.save(checkpoint_path)

    def clear_checkpoint(self, forum: str):
        """Clear checkpoint for a forum."""
        checkpoint_path = self.get_checkpoint_path(forum)
        if os.path.exists(checkpoint_path):
            os.remove(checkpoint_path)


class AsyncForumScraper(ABC):
    """Base class for async forum scrapers."""

    def __init__(
        self,
        async_scraper: AsyncScraper,
        db=None,
        output_dir: str = "./data"
    ):
        self.scraper = async_scraper
        self.db = db
        self.output_dir = output_dir
        self.checkpoint: Optional[ScrapeCheckpoint] = None

    @property
    @abstractmethod
    def forum_name(self) -> str:
        pass

    @property
    @abstractmethod
    def base_url(self) -> str:
        pass

    @abstractmethod
    def get_forum_sections(self) -> List[Dict[str, str]]:
        pass

    @abstractmethod
    def parse_thread_list(self, html: str, section_url: str) -> List[Dict[str, str]]:
        """Parse thread list from HTML."""
        pass

    @abstractmethod
    def parse_thread(self, html: str, url: str, category: str) -> Optional[Dict[str, Any]]:
        """Parse a thread from HTML."""
        pass

    def get_section_page_url(self, section_url: str, page: int) -> str:
        """Get URL for a section page."""
        if page == 1:
            return section_url
        return f"{section_url}page-{page}"

    async def scrape(
        self,
        max_pages: int = 5,
        max_threads_per_section: int = 50,
        resume: bool = True,
        save_interval: int = 10
    ) -> List[Dict[str, Any]]:
        """Main async scraping method."""
        # Load or create checkpoint
        if resume:
            self.checkpoint = self.scraper.load_or_create_checkpoint(self.forum_name)
        else:
            self.checkpoint = self.scraper.create_checkpoint(self.forum_name)

        sections = self.get_forum_sections()
        all_discussions = []
        scraped_count = 0

        try:
            for section in sections:
                section_name = section.get('name', 'Unknown')
                section_url = section.get('url', '')

                # Skip completed sections
                if section_name in self.checkpoint.completed_sections:
                    logger.info(f"Skipping completed section: {section_name}")
                    continue

                logger.info(f"Scraping section: {section_name}")
                self.checkpoint.current_section = section_name
                threads_in_section = 0

                # Determine starting page
                start_page = 1
                if self.checkpoint.current_section == section_name:
                    start_page = self.checkpoint.current_page

                for page in range(start_page, max_pages + 1):
                    if threads_in_section >= max_threads_per_section:
                        break

                    self.checkpoint.current_page = page
                    page_url = self.get_section_page_url(section_url, page)

                    # Fetch thread list
                    html = await self.scraper.fetch(page_url)
                    if not html:
                        self.checkpoint.stats['errors'] = self.checkpoint.stats.get('errors', 0) + 1
                        continue

                    threads = self.parse_thread_list(html, section_url)
                    if not threads:
                        break

                    # Fetch threads concurrently
                    urls_to_fetch = []
                    url_info = {}

                    for thread_info in threads:
                        if threads_in_section >= max_threads_per_section:
                            break

                        thread_url = thread_info.get('url', '')
                        if not thread_url or thread_url in self.checkpoint.completed_urls:
                            continue

                        urls_to_fetch.append(thread_url)
                        url_info[thread_url] = thread_info

                    # Fetch all thread pages
                    thread_htmls = await self.scraper.fetch_many(urls_to_fetch)

                    # Parse threads
                    for url, html in thread_htmls.items():
                        if html:
                            discussion = self.parse_thread(html, url, section_name)
                            if discussion:
                                all_discussions.append(discussion)
                                self.checkpoint.completed_urls.add(url)
                                threads_in_section += 1
                                scraped_count += 1

                                self.checkpoint.stats['threads_found'] = \
                                    self.checkpoint.stats.get('threads_found', 0) + 1

                                # Save to database if available
                                if self.db:
                                    self._save_to_db(discussion)

                        else:
                            # Track failed URLs for retry
                            retry_count = self.checkpoint.failed_urls.get(url, 0) + 1
                            self.checkpoint.failed_urls[url] = retry_count

                    # Periodic checkpoint save
                    if scraped_count % save_interval == 0:
                        self.scraper.save_checkpoint(self.checkpoint)
                        logger.info(f"Checkpoint saved: {scraped_count} threads scraped")

                # Mark section as completed
                self.checkpoint.completed_sections.append(section_name)
                self.checkpoint.current_page = 1
                self.scraper.save_checkpoint(self.checkpoint)

        except Exception as e:
            logger.error(f"Error during scraping: {e}")
            self.scraper.save_checkpoint(self.checkpoint)
            raise

        finally:
            # Final checkpoint save
            self.scraper.save_checkpoint(self.checkpoint)

        # Clear checkpoint on successful completion
        if not self.checkpoint.failed_urls:
            self.scraper.clear_checkpoint(self.forum_name)
            logger.info("Scraping completed successfully, checkpoint cleared")

        return all_discussions

    async def retry_failed(self) -> List[Dict[str, Any]]:
        """Retry previously failed URLs."""
        if not self.checkpoint:
            self.checkpoint = self.scraper.load_or_create_checkpoint(self.forum_name)

        if not self.checkpoint.failed_urls:
            logger.info("No failed URLs to retry")
            return []

        discussions = []
        urls_to_retry = [
            url for url, count in self.checkpoint.failed_urls.items()
            if count < self.scraper.max_retries
        ]

        logger.info(f"Retrying {len(urls_to_retry)} failed URLs")

        thread_htmls = await self.scraper.fetch_many(urls_to_retry)

        for url, html in thread_htmls.items():
            if html:
                discussion = self.parse_thread(html, url, 'Retry')
                if discussion:
                    discussions.append(discussion)
                    self.checkpoint.completed_urls.add(url)
                    del self.checkpoint.failed_urls[url]

                    if self.db:
                        self._save_to_db(discussion)

        self.scraper.save_checkpoint(self.checkpoint)
        return discussions

    def _save_to_db(self, discussion: Dict[str, Any]):
        """Save discussion to database."""
        if self.db:
            try:
                self.db.upsert_discussion(discussion)
            except Exception as e:
                logger.error(f"Error saving to database: {e}")


class AsyncMrExcelScraper(AsyncForumScraper):
    """Async scraper for MrExcel.com."""

    @property
    def forum_name(self) -> str:
        return "MrExcel"

    @property
    def base_url(self) -> str:
        return "https://www.mrexcel.com/board/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        return [
            {'name': 'Excel Questions', 'url': urljoin(self.base_url, 'forums/excel-questions.teknedia/')},
            {'name': 'Power Query', 'url': urljoin(self.base_url, 'forums/power-query.teknedia/')},
            {'name': 'Power Pivot', 'url': urljoin(self.base_url, 'forums/power-pivot.teknedia/')},
            {'name': 'Charts and Graphs', 'url': urljoin(self.base_url, 'forums/charts-and-charting-in-excel.teknedia/')},
        ]

    def parse_thread_list(self, html: str, section_url: str) -> List[Dict[str, str]]:
        soup = BeautifulSoup(html, 'lxml')
        threads = []

        for elem in soup.select('.structItem--thread'):
            title_elem = elem.select_one('.structItem-title a')
            time_elem = elem.select_one('.structItem-latestDate time')

            if title_elem:
                thread_url = urljoin(self.base_url, title_elem.get('href', ''))
                threads.append({
                    'title': title_elem.get_text(strip=True),
                    'url': thread_url,
                    'last_activity': time_elem.get('datetime', '') if time_elem else ''
                })

        return threads

    def parse_thread(self, html: str, url: str, category: str) -> Optional[Dict[str, Any]]:
        soup = BeautifulSoup(html, 'lxml')

        try:
            title_elem = soup.select_one('.p-title-value')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            author_elem = soup.select_one('.message--post .message-userDetails a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            date_elem = soup.select_one('.message--post time')
            date_created = date_elem.get('datetime', '') if date_elem else ""

            # Get posts
            posts = []
            for idx, post_elem in enumerate(soup.select('.message--post')[:50], 1):
                post_author = post_elem.select_one('.message-userDetails a')
                post_content = post_elem.select_one('.message-body')
                post_date = post_elem.select_one('time')

                posts.append({
                    'author': post_author.get_text(strip=True) if post_author else "Unknown",
                    'content': post_content.get_text(strip=True)[:10000] if post_content else "",
                    'date': post_date.get('datetime', '') if post_date else "",
                    'post_number': idx,
                    'is_solution': bool(post_elem.select_one('.label--solution'))
                })

            return {
                'title': title,
                'url': url,
                'author': author,
                'date_created': date_created,
                'category': category,
                'posts': posts,
                'forum_source': self.forum_name,
                'is_solved': bool(soup.select_one('.label--solved, .label--answered'))
            }

        except Exception as e:
            logger.error(f"Error parsing thread {url}: {e}")
            return None


async def run_async_scrape(
    forums: List[str],
    db=None,
    max_pages: int = 5,
    max_threads: int = 50,
    concurrent: int = 10,
    rate: float = 2.0,
    proxy_file: str = None,
    resume: bool = True
) -> Dict[str, List[Dict[str, Any]]]:
    """Run async scraping for multiple forums."""
    scraper = AsyncScraper(
        max_concurrent=concurrent,
        requests_per_second=rate,
        proxy_file=proxy_file
    )

    results = {}

    forum_scrapers = {
        'mrexcel': AsyncMrExcelScraper,
        # Add more async scrapers here
    }

    try:
        for forum in forums:
            if forum.lower() in forum_scrapers:
                scraper_class = forum_scrapers[forum.lower()]
                forum_scraper = scraper_class(scraper, db)

                logger.info(f"Starting async scrape of {forum}")
                discussions = await forum_scraper.scrape(
                    max_pages=max_pages,
                    max_threads_per_section=max_threads,
                    resume=resume
                )

                results[forum] = discussions
                logger.info(f"Completed {forum}: {len(discussions)} discussions")

    finally:
        await scraper.close()

    return results


# CLI integration
def main():
    import argparse

    parser = argparse.ArgumentParser(description='Async Excel Forum Scraper')
    parser.add_argument('--forum', choices=['mrexcel', 'all'], default='mrexcel')
    parser.add_argument('--pages', type=int, default=5)
    parser.add_argument('--threads', type=int, default=50)
    parser.add_argument('--concurrent', type=int, default=10)
    parser.add_argument('--rate', type=float, default=2.0)
    parser.add_argument('--proxy-file', help='File with proxy list')
    parser.add_argument('--no-resume', action='store_true', help='Start fresh, ignore checkpoints')

    args = parser.parse_args()

    forums = ['mrexcel'] if args.forum != 'all' else ['mrexcel']

    results = asyncio.run(run_async_scrape(
        forums=forums,
        max_pages=args.pages,
        max_threads=args.threads,
        concurrent=args.concurrent,
        rate=args.rate,
        proxy_file=args.proxy_file,
        resume=not args.no_resume
    ))

    total = sum(len(d) for d in results.values())
    print(f"\nTotal scraped: {total} discussions")


if __name__ == '__main__':
    main()
