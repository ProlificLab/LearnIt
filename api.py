"""
Flask REST API for Excel Forum Scraper.

Provides endpoints for:
- Browsing and searching discussions
- Managing scrape jobs
- Viewing enrichment data
- User profiles
- Statistics and analytics
"""

import os
import asyncio
import threading
from functools import wraps
from datetime import datetime
from flask import Flask, request, jsonify, Response
from flask_cors import CORS

from database import ForumDatabase
from enrichment import EnrichmentPipeline, CodeExtractor, DuplicateDetector, QualityScorer

app = Flask(__name__)
CORS(app)

# Configuration
DATABASE_PATH = os.environ.get('FORUM_DB_PATH', './data/forum_data.db')
OUTPUT_DIR = os.environ.get('FORUM_OUTPUT_DIR', './data')

# Initialize components
db = ForumDatabase(DATABASE_PATH)
enrichment_pipeline = EnrichmentPipeline()
code_extractor = CodeExtractor()
quality_scorer = QualityScorer()
duplicate_detector = DuplicateDetector()

# Background job tracking
background_jobs = {}


def json_response(data, status=200):
    """Create a JSON response."""
    return jsonify(data), status


def error_response(message, status=400):
    """Create an error response."""
    return jsonify({'error': message}), status


def paginate(query_func):
    """Decorator to add pagination to endpoints."""
    @wraps(query_func)
    def wrapper(*args, **kwargs):
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 20, type=int)
        per_page = min(per_page, 100)  # Max 100 per page

        offset = (page - 1) * per_page
        kwargs['limit'] = per_page
        kwargs['offset'] = offset

        results = query_func(*args, **kwargs)

        return {
            'data': results,
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total_returned': len(results)
            }
        }
    return wrapper


# ============== Health & Status ==============

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return json_response({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat()
    })


@app.route('/api/stats', methods=['GET'])
def get_statistics():
    """Get database statistics."""
    stats = db.get_statistics()
    formula_stats = db.get_formula_statistics()

    return json_response({
        'database': stats,
        'enrichment': formula_stats
    })


# ============== Discussions ==============

@app.route('/api/discussions', methods=['GET'])
def list_discussions():
    """List discussions with filtering and pagination."""
    forum = request.args.get('forum')
    category = request.args.get('category')
    order_by = request.args.get('order_by', 'last_updated')
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)

    per_page = min(per_page, 100)
    offset = (page - 1) * per_page

    # Validate order_by
    valid_orders = ['last_updated', 'date_created', 'views', 'replies']
    if order_by not in valid_orders:
        order_by = 'last_updated'

    discussions = db.get_discussions(
        forum_source=forum,
        category=category,
        limit=per_page,
        offset=offset,
        order_by=order_by
    )

    return json_response({
        'data': discussions,
        'pagination': {
            'page': page,
            'per_page': per_page,
            'total_returned': len(discussions)
        }
    })


@app.route('/api/discussions/<int:discussion_id>', methods=['GET'])
def get_discussion(discussion_id):
    """Get a single discussion with posts."""
    include_enrichment = request.args.get('enrichment', 'false').lower() == 'true'

    discussion = db.get_discussion_with_posts(discussion_id)
    if not discussion:
        return error_response('Discussion not found', 404)

    if include_enrichment:
        discussion['enrichment'] = db.get_enrichment(discussion_id)
        discussion['extracted_code'] = db.get_extracted_code(discussion_id)
        discussion['extracted_formulas'] = db.get_extracted_formulas(discussion_id)

    return json_response(discussion)


@app.route('/api/discussions/<int:discussion_id>/history', methods=['GET'])
def get_discussion_history(discussion_id):
    """Get historical data for a discussion."""
    history = db.get_discussion_history(discussion_id)
    return json_response({'history': history})


@app.route('/api/discussions/<int:discussion_id>/duplicates', methods=['GET'])
def get_discussion_duplicates(discussion_id):
    """Get discussions marked as duplicates."""
    duplicates = db.get_duplicates(discussion_id)
    return json_response({'duplicates': duplicates})


# ============== Search ==============

@app.route('/api/search', methods=['GET'])
def search_discussions():
    """Search discussions."""
    query = request.args.get('q', '')
    limit = request.args.get('limit', 50, type=int)

    if not query:
        return error_response('Query parameter "q" is required')

    results = db.search_discussions(query, limit=min(limit, 100))
    return json_response({
        'query': query,
        'results': results,
        'count': len(results)
    })


@app.route('/api/search/code', methods=['GET'])
def search_by_code():
    """Search discussions containing code."""
    code_type = request.args.get('type')  # 'vba', 'formula', 'sql'
    limit = request.args.get('limit', 50, type=int)

    results = db.get_discussions_with_code(code_type, limit=min(limit, 100))
    return json_response({
        'code_type': code_type,
        'results': results,
        'count': len(results)
    })


# ============== Users ==============

@app.route('/api/users', methods=['GET'])
def list_users():
    """List user profiles."""
    forum = request.args.get('forum')
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)

    per_page = min(per_page, 100)
    offset = (page - 1) * per_page

    users = db.get_users(forum_source=forum, limit=per_page, offset=offset)

    return json_response({
        'data': users,
        'pagination': {
            'page': page,
            'per_page': per_page,
            'total_returned': len(users)
        }
    })


@app.route('/api/users/<username>', methods=['GET'])
def get_user(username):
    """Get a user profile."""
    forum = request.args.get('forum')

    user = db.get_user(username, forum) if forum else None

    if not user:
        # Try to find in any forum
        for f in ['MrExcel', 'ExcelForum', 'Chandoo']:
            user = db.get_user(username, f)
            if user:
                break

    if not user:
        return error_response('User not found', 404)

    return json_response(user)


# ============== Watched Threads ==============

@app.route('/api/watched', methods=['GET'])
def list_watched_threads():
    """List watched threads."""
    watched = db.get_watched_threads()
    return json_response({'watched': watched})


@app.route('/api/watched', methods=['POST'])
def add_watched_thread():
    """Add a thread to watchlist."""
    data = request.get_json()

    if not data or 'url' not in data:
        return error_response('URL is required')

    url = data['url']
    frequency = data.get('frequency', 'daily')
    notes = data.get('notes', '')

    # Find discussion by URL
    discussion = db.get_discussion_by_url(url)
    if not discussion:
        return error_response('Discussion not found. Scrape it first.', 404)

    db.watch_thread(discussion['id'], frequency, notes)
    return json_response({'message': 'Thread added to watchlist', 'id': discussion['id']})


@app.route('/api/watched/<int:discussion_id>', methods=['DELETE'])
def remove_watched_thread(discussion_id):
    """Remove a thread from watchlist."""
    success = db.unwatch_thread(discussion_id)
    if success:
        return json_response({'message': 'Thread removed from watchlist'})
    return error_response('Thread not found in watchlist', 404)


# ============== Enrichment ==============

@app.route('/api/enrich/<int:discussion_id>', methods=['POST'])
def enrich_discussion(discussion_id):
    """Enrich a single discussion."""
    discussion = db.get_discussion_with_posts(discussion_id)
    if not discussion:
        return error_response('Discussion not found', 404)

    # Run enrichment
    enriched = enrichment_pipeline.enrich_discussion(discussion)
    enrichment_data = enriched.get('enrichment', {})

    # Save to database
    db.save_enrichment(discussion_id, enrichment_data)

    # Save extracted code and formulas
    for code in enrichment_data.get('vba_code', []):
        db.save_extracted_code(discussion_id, 1, {
            'code_type': 'vba',
            'content': code.get('content', ''),
            'functions_used': code.get('functions', []),
            'line_count': code.get('content', '').count('\n') + 1
        })

    for formula in enrichment_data.get('formulas', []):
        db.save_extracted_formula(discussion_id, 1, formula)

    return json_response({
        'message': 'Discussion enriched',
        'enrichment': enrichment_data
    })


@app.route('/api/enrich/batch', methods=['POST'])
def enrich_batch():
    """Enrich multiple discussions (background job)."""
    data = request.get_json()
    forum = data.get('forum') if data else None
    limit = data.get('limit', 100) if data else 100

    job_id = f"enrich_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    def run_enrichment():
        discussions = db.get_discussions(forum_source=forum, limit=limit)
        total = len(discussions)
        processed = 0

        background_jobs[job_id] = {
            'status': 'running',
            'total': total,
            'processed': 0,
            'started_at': datetime.now().isoformat()
        }

        for disc in discussions:
            full_disc = db.get_discussion_with_posts(disc['id'])
            if full_disc:
                enriched = enrichment_pipeline.enrich_discussion(full_disc)
                enrichment_data = enriched.get('enrichment', {})
                db.save_enrichment(disc['id'], enrichment_data)

            processed += 1
            background_jobs[job_id]['processed'] = processed

        background_jobs[job_id]['status'] = 'completed'
        background_jobs[job_id]['completed_at'] = datetime.now().isoformat()

    thread = threading.Thread(target=run_enrichment)
    thread.start()

    return json_response({
        'job_id': job_id,
        'message': 'Enrichment job started'
    })


@app.route('/api/analyze/formula', methods=['POST'])
def analyze_formula():
    """Analyze an Excel formula."""
    data = request.get_json()
    formula = data.get('formula', '') if data else ''

    if not formula:
        return error_response('Formula is required')

    if not formula.startswith('='):
        formula = '=' + formula

    extracted = code_extractor.extract_formulas(formula)

    if extracted:
        result = {
            'formula': extracted[0].formula,
            'functions': extracted[0].functions,
            'cell_references': extracted[0].cell_references,
            'complexity_score': extracted[0].complexity_score,
            'is_array_formula': extracted[0].is_array_formula
        }
    else:
        result = {
            'formula': formula,
            'functions': [],
            'error': 'Could not parse formula'
        }

    return json_response(result)


@app.route('/api/analyze/code', methods=['POST'])
def analyze_code():
    """Analyze VBA code."""
    data = request.get_json()
    code = data.get('code', '') if data else ''

    if not code:
        return error_response('Code is required')

    extracted = code_extractor.extract_all(code)

    return json_response({
        'vba_snippets': len(extracted['vba_code']),
        'formulas': len(extracted['formulas']),
        'sql_snippets': len(extracted['sql_code']),
        'vba_code': [
            {
                'content': c.content[:500] + '...' if len(c.content) > 500 else c.content,
                'functions': c.functions_used,
                'line_count': c.line_count
            }
            for c in extracted['vba_code']
        ],
        'formulas': [
            {
                'formula': f.formula,
                'functions': f.functions,
                'complexity': f.complexity_score
            }
            for f in extracted['formulas']
        ]
    })


@app.route('/api/analyze/similarity', methods=['POST'])
def check_similarity():
    """Check similarity between two texts."""
    data = request.get_json()
    text1 = data.get('text1', '') if data else ''
    text2 = data.get('text2', '') if data else ''

    if not text1 or not text2:
        return error_response('Both text1 and text2 are required')

    similarity = duplicate_detector.compute_similarity(text1, text2)

    return json_response({
        'similarity': round(similarity, 4),
        'is_duplicate': similarity >= 0.7
    })


# ============== Jobs ==============

@app.route('/api/jobs', methods=['GET'])
def list_jobs():
    """List background jobs."""
    return json_response({'jobs': background_jobs})


@app.route('/api/jobs/<job_id>', methods=['GET'])
def get_job_status(job_id):
    """Get status of a background job."""
    if job_id not in background_jobs:
        return error_response('Job not found', 404)

    return json_response(background_jobs[job_id])


# ============== Community Sources ==============

@app.route('/api/community/sources', methods=['GET'])
def list_community_sources():
    """List available community sources."""
    return json_response({
        'sources': [
            {
                'id': 'stackoverflow',
                'name': 'Stack Overflow',
                'description': 'Q&A for Excel, VBA, Power BI, DAX',
                'tags': ['excel', 'vba', 'powerbi', 'dax', 'power-query']
            },
            {
                'id': 'reddit',
                'name': 'Reddit',
                'description': 'Community discussions from r/excel, r/vba, r/PowerBI',
                'subreddits': ['excel', 'vba', 'PowerBI', 'excelevator']
            },
            {
                'id': 'mstech',
                'name': 'Microsoft Tech Community',
                'description': 'Official Microsoft community forums',
                'boards': ['excel', 'powerbi', 'powerquery']
            },
            {
                'id': 'powerbi',
                'name': 'Power BI Community',
                'description': 'Official Power BI community',
                'categories': ['desktop', 'service', 'dax', 'dataflows']
            },
            {
                'id': 'superuser',
                'name': 'Super User',
                'description': 'Stack Exchange site for power users and Excel questions',
                'tags': ['microsoft-excel', 'excel-formula', 'spreadsheet', 'google-sheets']
            },
            {
                'id': 'youtube',
                'name': 'YouTube',
                'description': 'Excel/Power BI tutorial videos and comments',
                'requires_api_key': True,
                'queries': ['excel tutorial', 'power bi tutorial', 'vba tutorial']
            },
            {
                'id': 'github',
                'name': 'GitHub Issues',
                'description': 'Issues from Excel/Power BI related repositories',
                'repos': ['openpyxl/openpyxl', 'SheetJS/sheetjs', 'microsoft/powerbi-client-python']
            },
            {
                'id': 'ozgrid',
                'name': 'OzGrid Forum',
                'description': 'Long-standing Excel help forum',
                'categories': ['excel-general', 'excel-vba', 'excel-formulas']
            },
            {
                'id': 'quora',
                'name': 'Quora',
                'description': 'Q&A platform with Excel and Power BI topics',
                'topics': ['Microsoft-Excel', 'Visual-Basic-for-Applications-VBA', 'Microsoft-Power-BI']
            }
        ]
    })


@app.route('/api/community/scrape', methods=['POST'])
def scrape_community():
    """Start a community scraping job (background)."""
    data = request.get_json() or {}
    source = data.get('source', 'all')
    pages = data.get('pages', 5)
    youtube_api_key = data.get('youtube_api_key')
    github_token = data.get('github_token')

    job_id = f"community_{source}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    def run_community_scrape():
        from community_scrapers import (
            StackOverflowScraper, RedditScraper,
            MSTechCommunityScraper, PowerBICommunityScraper,
            SuperUserScraper, YouTubeScraper, GitHubIssuesScraper,
            OzGridScraper, QuoraScraper,
            TopicAnalyzer
        )

        # Define all scrapers with their initialization
        def get_scrapers():
            return {
                'stackoverflow': lambda: StackOverflowScraper(db),
                'reddit': lambda: RedditScraper(db),
                'mstech': lambda: MSTechCommunityScraper(db),
                'powerbi': lambda: PowerBICommunityScraper(db),
                'superuser': lambda: SuperUserScraper(db),
                'youtube': lambda: YouTubeScraper(db, api_key=youtube_api_key),
                'github': lambda: GitHubIssuesScraper(db, token=github_token),
                'ozgrid': lambda: OzGridScraper(db),
                'quora': lambda: QuoraScraper(db),
            }

        scrapers = get_scrapers()
        sources_to_run = list(scrapers.keys()) if source == 'all' else [source]
        all_questions = []

        background_jobs[job_id] = {
            'status': 'running',
            'source': source,
            'started_at': datetime.now().isoformat(),
            'scraped': 0,
            'sources_completed': []
        }

        for src in sources_to_run:
            if src in scrapers:
                try:
                    scraper = scrapers[src]()

                    # Different scrape parameters for different sources
                    if src == 'youtube':
                        questions = scraper.scrape(max_results=pages * 10)
                    elif src == 'github':
                        questions = scraper.scrape(per_repo=pages * 20)
                    else:
                        questions = scraper.scrape(pages=pages)

                    all_questions.extend(questions)
                    background_jobs[job_id]['scraped'] = len(all_questions)
                    background_jobs[job_id]['sources_completed'].append(src)
                except Exception as e:
                    if 'errors' not in background_jobs[job_id]:
                        background_jobs[job_id]['errors'] = []
                    background_jobs[job_id]['errors'].append(f"{src}: {str(e)}")

        # Run topic analysis
        if all_questions:
            analyzer = TopicAnalyzer(all_questions)
            background_jobs[job_id]['insights'] = {
                'common_topics': analyzer.get_common_words(20),
                'unanswered': analyzer.get_unanswered_topics(10),
                'by_source': analyzer.get_source_summary()
            }

        background_jobs[job_id]['status'] = 'completed'
        background_jobs[job_id]['completed_at'] = datetime.now().isoformat()

    thread = threading.Thread(target=run_community_scrape)
    thread.start()

    return json_response({
        'job_id': job_id,
        'message': f'Community scraping started for {source}'
    })


@app.route('/api/community/trending', methods=['GET'])
def get_trending_topics():
    """Get trending topics from communities."""
    source = request.args.get('source', 'reddit')

    try:
        if source == 'reddit':
            from community_scrapers import RedditScraper
            scraper = RedditScraper()
            topics = scraper.get_trending_topics('excel', limit=20)
            return json_response({'source': 'reddit', 'topics': topics})

        elif source == 'stackoverflow':
            from community_scrapers import StackOverflowScraper
            scraper = StackOverflowScraper()
            tags = scraper.get_trending_tags(count=30)
            return json_response({'source': 'stackoverflow', 'trending_tags': tags})

        elif source == 'powerbi':
            from community_scrapers import PowerBICommunityScraper
            scraper = PowerBICommunityScraper()
            ideas = scraper.get_ideas(pages=2)
            return json_response({'source': 'powerbi', 'feature_ideas': ideas})

        elif source == 'github':
            from community_scrapers import GitHubIssuesScraper
            scraper = GitHubIssuesScraper()
            # Search for trending Excel-related issues
            issues = scraper.search_issues('excel', max_results=30)
            return json_response({
                'source': 'github',
                'trending_issues': [
                    {'title': q.title, 'url': q.url, 'score': q.score, 'comments': q.answers_count}
                    for q in issues
                ]
            })

        else:
            return error_response(f'Unknown source: {source}. Available: reddit, stackoverflow, powerbi, github')

    except Exception as e:
        return error_response(f'Error fetching trends: {str(e)}', 500)


@app.route('/api/community/analyze', methods=['GET'])
def analyze_community_data():
    """Analyze scraped community data."""
    # Get all discussions from community sources
    community_sources = [
        'StackOverflow', 'Reddit', 'MSTechCommunity', 'PowerBICommunity',
        'SuperUser', 'YouTube', 'GitHub', 'OzGrid', 'Quora'
    ]

    discussions = []
    for source in community_sources:
        # Also check for sources with slashes (like GitHub/repo or Reddit/r/sub)
        discussions.extend(db.get_discussions(forum_source=source, limit=1000))

    if not discussions:
        return json_response({
            'message': 'No community data found. Run /api/community/scrape first.',
            'total': 0
        })

    # Analyze
    from community_scrapers import Question, TopicAnalyzer

    questions = []
    for d in discussions:
        q = Question(
            title=d.get('title', ''),
            url=d.get('url', ''),
            body='',
            author=d.get('author', ''),
            created_at=d.get('date_created', ''),
            tags=d.get('tags', '').split(',') if isinstance(d.get('tags'), str) else [],
            views=d.get('views', 0),
            answers_count=d.get('replies', 0),
            is_answered=bool(d.get('is_solved')),
            source=d.get('forum_source', ''),
            category=d.get('category', '')
        )
        questions.append(q)

    analyzer = TopicAnalyzer(questions)

    return json_response({
        'total_questions': len(questions),
        'common_topics': analyzer.get_common_words(30),
        'unanswered_topics': analyzer.get_unanswered_topics(20),
        'tag_distribution': dict(list(analyzer.get_tag_distribution().items())[:30]),
        'by_source': analyzer.get_source_summary()
    })


# ============== Scraping ==============

@app.route('/api/scrape', methods=['POST'])
def start_scrape():
    """Start a scraping job (background)."""
    data = request.get_json() or {}
    forum = data.get('forum', 'all')
    pages = data.get('pages', 5)
    threads = data.get('threads', 50)
    incremental = data.get('incremental', True)

    job_id = f"scrape_{forum}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    def run_scrape():
        from excel_forum_scraper import get_scraper

        forums = ['mrexcel', 'excelforum', 'chandoo'] if forum == 'all' else [forum]
        total_scraped = 0

        background_jobs[job_id] = {
            'status': 'running',
            'forum': forum,
            'started_at': datetime.now().isoformat(),
            'scraped': 0
        }

        for f in forums:
            scraper = get_scraper(f, db, OUTPUT_DIR, 1.0)
            if scraper:
                try:
                    discussions = scraper.scrape(
                        max_pages=pages,
                        max_threads_per_section=threads,
                        incremental=incremental
                    )
                    total_scraped += len(discussions)
                    background_jobs[job_id]['scraped'] = total_scraped
                except Exception as e:
                    background_jobs[job_id]['error'] = str(e)

        background_jobs[job_id]['status'] = 'completed'
        background_jobs[job_id]['completed_at'] = datetime.now().isoformat()

    thread = threading.Thread(target=run_scrape)
    thread.start()

    return json_response({
        'job_id': job_id,
        'message': f'Scraping job started for {forum}'
    })


# ============== Export ==============

@app.route('/api/export', methods=['GET'])
def export_data():
    """Export data as JSON."""
    forum = request.args.get('forum')

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"export_{timestamp}.json"
    filepath = os.path.join(OUTPUT_DIR, filename)

    count = db.export_to_json(filepath, forum)

    return json_response({
        'message': f'Exported {count} discussions',
        'file': filepath
    })


# ============== Formula Statistics ==============

@app.route('/api/formulas/stats', methods=['GET'])
def formula_statistics():
    """Get formula usage statistics."""
    stats = db.get_formula_statistics()
    return json_response(stats)


@app.route('/api/formulas/top-functions', methods=['GET'])
def top_functions():
    """Get most used Excel functions."""
    limit = request.args.get('limit', 20, type=int)
    stats = db.get_formula_statistics()

    top = dict(list(stats.get('top_functions', {}).items())[:limit])
    return json_response({'functions': top})


# ============== Error Handlers ==============

@app.errorhandler(404)
def not_found(e):
    return error_response('Endpoint not found', 404)


@app.errorhandler(500)
def server_error(e):
    return error_response('Internal server error', 500)


# ============== Main ==============

def create_app(database_path=None):
    """Application factory."""
    global db

    if database_path:
        db = ForumDatabase(database_path)

    return app


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Excel Forum Scraper API')
    parser.add_argument('--host', default='0.0.0.0', help='Host to bind to')
    parser.add_argument('--port', type=int, default=5000, help='Port to bind to')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    parser.add_argument('--database', default=DATABASE_PATH, help='Database path')

    args = parser.parse_args()

    db = ForumDatabase(args.database)

    print(f"""
    ╔══════════════════════════════════════════════════════════╗
    ║         Excel Forum Scraper API                          ║
    ╠══════════════════════════════════════════════════════════╣
    ║  Endpoints:                                              ║
    ║    GET  /api/health          - Health check              ║
    ║    GET  /api/stats           - Database statistics       ║
    ║    GET  /api/discussions     - List discussions          ║
    ║    GET  /api/search?q=       - Search discussions        ║
    ║    GET  /api/users           - List users                ║
    ║    POST /api/scrape          - Start scraping job        ║
    ║    POST /api/enrich/<id>     - Enrich discussion         ║
    ║    POST /api/analyze/formula - Analyze formula           ║
    ╚══════════════════════════════════════════════════════════╝
    """)

    app.run(host=args.host, port=args.port, debug=args.debug)
