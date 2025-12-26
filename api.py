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
