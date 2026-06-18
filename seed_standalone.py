import os, sys
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(__file__))
os.environ['DATABASE_URL'] = ''

# Bypass main.py auto-seed by importing only what we need
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker

DB_PATH = os.path.join(os.path.dirname(__file__), "db", "orders.db")
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
Base = declarative_base()

class Order(Base):
    __tablename__ = "orders"
    id = Column(Integer, primary_key=True, index=True)
    order_id = Column(String, unique=True, index=True)
    created_at = Column(DateTime, default=datetime.now)
    amount_usdt = Column(Float)
    amount_rub = Column(Float)
    rate_at_creation = Column(Float)
    commission_percent = Column(Float, default=3.0)
    commission_amount = Column(Float, default=0.0)
    currency = Column(String, default="RUB")
    bank = Column(String, default="")
    phone = Column(String, default="")
    deposit_address = Column(String, default="")
    status = Column(String, default="pending")
    order_type = Column(String, default="buy")
    asset_type = Column(String, default="USDT")
    wallet = Column(String, default="")

class ChatSession(Base):
    __tablename__ = "chat_sessions"
    id = Column(Integer, primary_key=True, index=True)
    client_name = Column(String, default="Клиент")
    email = Column(String, default="")
    status = Column(String, default="active")
    unread = Column(Integer, default=0)
    ip_address = Column(String(45), default="")
    country_code = Column(String(2), default="")
    country_name = Column(String(100), default="")
    wallet = Column(String(100), default="")
    created_at = Column(DateTime, default=datetime.now)

class ChatMessage(Base):
    __tablename__ = "chat_messages"
    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("chat_sessions.id"))
    sender = Column(String)
    message = Column(String)
    created_at = Column(DateTime, default=datetime.now)

class SupportTicket(Base):
    __tablename__ = "support_tickets"
    id = Column(Integer, primary_key=True, index=True)
    created_at = Column(DateTime, default=datetime.now)
    deposit_address = Column(String, nullable=True)
    order_id = Column(String, nullable=True)
    email = Column(String)
    message = Column(String)
    status = Column(String, default="pending")

Base.metadata.create_all(bind=engine)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

db = SessionLocal()

tickets = [
    SupportTicket(email="ivan@mail.com", message="Здравствуйте! Я отправил 50 USDT на указанный адрес 15 минут назад, но статус заказа #A1B2C3D4 всё ещё «Ожидает оплаты». Подскажите, что делать?", status="pending"),
    SupportTicket(email="maria@yandex.ru", message="Добрый день! Перевела 120 USDT, заказ #E5F6G7H8. Транзакция подтверждена в сети, но на сайте статус не обновляется. Проверьте пожалуйста.", status="pending"),
    SupportTicket(email="alex@bk.ru", message="Хотел бы узнать лимиты. Какая максимальная сумма одной операции? Планирую обменять 5000 USDT на рубли.", status="resolved"),
    SupportTicket(email="elena@mail.ru", message="Здравствуйте! По ошибке отправил USDT на старый адрес. Транзакция прошла, но заказ не создавался. Можно вернуть средства?", status="pending"),
    SupportTicket(email="sergey@gmail.com", message="Какие банки поддерживаются для вывода рублей? Интересует Сбербанк и Тинькофф. И какие минимальные суммы вывода?", status="resolved"),
    SupportTicket(email="dmitry@mail.ru", message="Здравствуйте! Заказ #N3O4P5Q6 оплачен более часа назад, но средства на карту ещё не поступили. Сколько обычно занимает перевод?", status="pending"),
]
for t in tickets:
    db.add(t)
db.commit()
print(f"Seeded {len(tickets)} support tickets")

orders = [
    Order(order_id="A1B2C3D4", created_at=datetime.now()-timedelta(minutes=45),
          amount_usdt=50.0, amount_rub=3668.54, rate_at_creation=75.64,
          commission_percent=3.0, commission_amount=113.46,
          currency="RUB", bank="Сбербанк", phone="+7 999 123 45 67",
          deposit_address="TR7NHqjeKQxGTCi8q8ZY4pL8otSzgjLj6t",
          status="pending", order_type="buy", asset_type="USDT",
          wallet="9x4K8J3p2QmR7vW1nL5tY6cB0fA2dE3gH6i"),
    Order(order_id="E5F6G7H8", created_at=datetime.now()-timedelta(hours=2),
          amount_usdt=120.0, amount_rub=1276651.92, rate_at_creation=10967.80,
          commission_percent=3.0, commission_amount=39484.08,
          currency="RUB", bank="Тинькофф", phone="+7 916 555 77 88",
          deposit_address="4K8J3p2QmR7vW1nL5tY6cB0fA2dE3gH6i",
          status="paid", order_type="buy", asset_type="SOL",
          wallet="5tY6cB0fA2dE3gH6i9x4K8J3p2QmR7vW1nL"),
    Order(order_id="J9K0L1M2", created_at=datetime.now()-timedelta(days=2),
          amount_usdt=250.0, amount_rub=64199450.00, rate_at_creation=264740.00,
          commission_percent=3.0, commission_amount=1985550.00,
          currency="RUB", bank="Сбербанк", phone="+7 903 222 33 44",
          deposit_address="0x4K8J3p2QmR7vW1nL5tY6cB0fA2dE3gH6i",
          status="canceled", order_type="buy", asset_type="ETH",
          wallet="0x9x4K8J3p2QmR7vW1nL5tY6cB0fA2dE3gH6i"),
    Order(order_id="N3O4P5Q6", created_at=datetime.now()-timedelta(hours=18),
          amount_usdt=100.0, amount_rub=7790.92, rate_at_creation=75.64,
          commission_percent=3.0, commission_amount=226.92,
          currency="RUB", bank="Альфа-Банк", phone="+7 985 444 55 66",
          deposit_address="TR7NHqjeKQxGTCi8q8ZY4pL8otSzgjLj6t",
          status="paid", order_type="sell", asset_type="USDT",
          wallet="F8dE3gH6i9x4K8J3p2QmR7vW1nL5tY6cB0"),
    Order(order_id="K2L3M4N5", created_at=datetime.now()-timedelta(hours=6),
          amount_usdt=75.0, amount_rub=5502.81, rate_at_creation=75.64,
          commission_percent=3.0, commission_amount=170.19,
          currency="RUB", bank="Тинькофф", phone="+7 911 888 99 00",
          deposit_address="B0fA2dE3gH6i9x4K8J3p2QmR7vW1nL5tY6",
          status="pending", order_type="buy", asset_type="USDT",
          wallet="W1nL5tY6cB0fA2dE3gH6i9x4K8J3p2QmR7v"),
    Order(order_id="P7R8S9T0", created_at=datetime.now()-timedelta(hours=48),
          amount_usdt=500.0, amount_rub=5648417.00, rate_at_creation=10967.80,
          commission_percent=3.0, commission_amount=164517.00,
          currency="RUB", bank="Сбербанк", phone="+7 926 111 22 33",
          deposit_address="SOLp2QmR7vW1nL5tY6cB0fA2dE3gH6i9x4",
          status="canceled", order_type="sell", asset_type="SOL",
          wallet="R7vW1nL5tY6cB0fA2dE3gH6i9x4K8J3p2Qm"),
]
for o in orders:
    db.add(o)
db.commit()
print(f"Seeded {len(orders)} orders")

sessions = [
    ChatSession(id=1, created_at=datetime.now()-timedelta(hours=3),
                status="active", unread=2,
                ip_address="195.122.210.10", country_code="RU", country_name="Russia",
                wallet="9x4K8J3p2QmR7vW1nL5tY6cB0fA2dE3gH6i"),
    ChatSession(id=2, created_at=datetime.now()-timedelta(days=1),
                status="closed", unread=0,
                ip_address="85.26.183.45", country_code="RU", country_name="Russia",
                wallet="5tY6cB0fA2dE3gH6i9x4K8J3p2QmR7vW1nL"),
    ChatSession(id=3, created_at=datetime.now()-timedelta(hours=12),
                status="active", unread=1,
                ip_address="176.59.12.89", country_code="RU", country_name="Russia",
                wallet="F8dE3gH6i9x4K8J3p2QmR7vW1nL5tY6cB0"),
]
for s in sessions:
    db.add(s)
db.commit()
print(f"Seeded {len(sessions)} chat sessions")

messages = [
    ChatMessage(session_id=1, sender="client",
                message="Здравствуйте! Отправил 50 USDT, статус не меняется уже 15 минут. Заказ #A1B2C3D4",
                created_at=datetime.now()-timedelta(hours=3)),
    ChatMessage(session_id=1, sender="admin",
                message="Здравствуйте! Проверяю ваш платёж. Подождите несколько минут, транзакция обрабатывается сетью",
                created_at=datetime.now()-timedelta(hours=2, minutes=55)),
    ChatMessage(session_id=1, sender="client",
                message="Спасибо, всё пришло! Заказ выполнен, деньги получил",
                created_at=datetime.now()-timedelta(hours=2, minutes=30)),
    ChatMessage(session_id=2, sender="client",
                message="Добрый день! Какие лимиты на одну операцию?",
                created_at=datetime.now()-timedelta(days=1)),
    ChatMessage(session_id=2, sender="admin",
                message="Добрый день! Минимальная сумма 10 USDT, максимальная 5000 USDT",
                created_at=datetime.now()-timedelta(days=1, hours=-1)),
    ChatMessage(session_id=2, sender="client",
                message="Понял, спасибо!",
                created_at=datetime.now()-timedelta(days=1, hours=-1, minutes=5)),
    ChatMessage(session_id=3, sender="client",
                message="Здравствуйте! Отправил 100 USDT по заказу #N3O4P5Q6, когда поступят рубли на карту?",
                created_at=datetime.now()-timedelta(hours=12)),
    ChatMessage(session_id=3, sender="admin",
                message="Здравствуйте! Ваш платёж получен, перевод на карту обычно занимает до 30 минут",
                created_at=datetime.now()-timedelta(hours=11, minutes=50)),
    ChatMessage(session_id=3, sender="client",
                message="Уже прошло 40 минут, денег всё нет",
                created_at=datetime.now()-timedelta(hours=11, minutes=20)),
]
for m in messages:
    db.add(m)
db.commit()
print(f"Seeded {len(messages)} chat messages")

db.close()
print("Done")
