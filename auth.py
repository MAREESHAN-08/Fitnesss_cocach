import sqlite3
import hashlib
import json
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fitness_database.db")
LEGACY_USERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")

def get_db_connection():
    """Establish and return a connection to the SQLite database."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def hash_password(password):
    """Generate SHA-256 hash of the password."""
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def init_db():
    """Initialize database tables and migrate legacy json users if applicable."""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Users table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # 2. Workout History table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workout_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            exercise_name TEXT NOT NULL,
            total_reps INTEGER DEFAULT 0,
            correct_reps INTEGER DEFAULT 0,
            wrong_reps INTEGER DEFAULT 0,
            form_score REAL DEFAULT 0,
            is_correct INTEGER DEFAULT 1,
            feedback TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    
    # Check if users need migration or default creation
    cursor.execute("SELECT COUNT(*) as count FROM users")
    user_count = cursor.fetchone()['count']
    
    if user_count == 0:
        migrated_any = False
        # Try migration from legacy users.json
        if os.path.exists(LEGACY_USERS_FILE):
            try:
                with open(LEGACY_USERS_FILE, "r") as f:
                    legacy_users = json.load(f)
                for uname, udata in legacy_users.items():
                    cursor.execute(
                        "INSERT OR IGNORE INTO users (username, name, password_hash) VALUES (?, ?, ?)",
                        (uname.lower().strip(), udata.get("name", uname), udata.get("password_hash"))
                    )
                conn.commit()
                migrated_any = True
            except Exception as e:
                print(f"Error migrating users.json: {e}")
                
        # If still empty, add default users
        if not migrated_any:
            default_users = [
                ("student", "Data Science Student", hash_password("fitness123")),
                ("admin", "Project Admin", hash_password("admin123"))
            ]
            cursor.executemany(
                "INSERT OR IGNORE INTO users (username, name, password_hash) VALUES (?, ?, ?)",
                default_users
            )
            conn.commit()
            
    conn.close()

# Auto-initialize database
init_db()

def verify_user(username, password):
    """Verify user login against SQLite database."""
    username = username.strip().lower()
    pw_hash = hash_password(password)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT username, name, password_hash FROM users WHERE username = ?",
        (username,)
    )
    user = cursor.fetchone()
    conn.close()
    
    if user and user['password_hash'] == pw_hash:
        return True, user['name']
    return False, None

def register_user(username, name, password):
    """Register a new user in the SQLite database."""
    username = username.strip().lower()
    name = name.strip()
    
    if not username or not password or not name:
        return False, "All fields are required!"
        
    if len(password) < 4:
        return False, "Password must be at least 4 characters long."
        
    pw_hash = hash_password(password)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO users (username, name, password_hash) VALUES (?, ?, ?)",
            (username, name, pw_hash)
        )
        conn.commit()
        return True, f"Account '{username}' registered in SQLite database! You can now log in."
    except sqlite3.IntegrityError:
        return False, f"Username '{username}' already exists. Please choose a different username."
    except Exception as e:
        return False, f"Database error: {str(e)}"
    finally:
        conn.close()

def log_workout_session(username, exercise_name, total_reps, correct_reps, wrong_reps, form_score, is_correct, feedback=""):
    """Saves a workout audit record into SQLite database."""
    if not username:
        username = "guest"
    username = username.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO workout_history 
            (username, exercise_name, total_reps, correct_reps, wrong_reps, form_score, is_correct, feedback, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            username,
            exercise_name.title(),
            int(total_reps),
            int(correct_reps),
            int(wrong_reps),
            float(form_score),
            1 if is_correct else 0,
            feedback,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        conn.commit()
        return True
    except Exception as e:
        print(f"Error logging workout: {e}")
        return False
    finally:
        conn.close()

def get_user_workout_history(username, limit=15):
    """Fetch recent workout history for a user from SQLite."""
    if not username:
        return []
    username = username.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, exercise_name, total_reps, correct_reps, wrong_reps, form_score, is_correct, feedback, created_at
        FROM workout_history
        WHERE username = ?
        ORDER BY id DESC
        LIMIT ?
    """, (username, limit))
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_user_stats(username):
    """Calculate aggregate stats for user from SQLite."""
    if not username:
        return {'total_sessions': 0, 'grand_total_reps': 0, 'grand_correct_reps': 0, 'avg_score': 0}
    username = username.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT 
            COUNT(*) as total_sessions,
            COALESCE(SUM(total_reps), 0) as grand_total_reps,
            COALESCE(SUM(correct_reps), 0) as grand_correct_reps,
            COALESCE(AVG(form_score), 0) as avg_score
        FROM workout_history
        WHERE username = ?
    """, (username,))
    row = cursor.fetchone()
    stats = dict(row) if row else {'total_sessions': 0, 'grand_total_reps': 0, 'grand_correct_reps': 0, 'avg_score': 0}
    conn.close()
    return stats
