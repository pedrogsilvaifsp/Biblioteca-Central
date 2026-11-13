from datetime import date
from flask_login import current_user
from sqlalchemy import and_

def inject_globals():
    unread_notifications = 0
    unread_requests = 0
    if current_user.is_authenticated:
        # Imports locais evitam ciclos durante a inicialização do Flask.
        from .models import Notification, NotificationRead, WorkRequest, WorkRequestRead
        from .extensions import db
        unread_notifications = Notification.query.filter(
            Notification.resolved_at.is_(None),
            ~Notification.reads.any(NotificationRead.user_id == current_user.id),
        ).count()
        unread_requests = WorkRequest.query.filter(
            ~WorkRequest.reads.any(
                and_(WorkRequestRead.user_id == current_user.id,
                     WorkRequestRead.last_viewed_at >= WorkRequest.updated_at)
            )
        ).count()
    return {
        "now_date": date.today().strftime("%d/%m/%Y"),
        "today": date.today(),
        "date": date,
        "timedelta": __import__("datetime").timedelta,
        "unread_notifications_count": unread_notifications,
        "unread_requests_count": unread_requests,
    }
