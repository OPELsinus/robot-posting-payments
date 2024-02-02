import datetime
from time import sleep

from openpyxl import load_workbook

from config import logger, engine_kwargs, project_name, smtp_host, smtp_author
from tools.smtp import smtp_send

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, Float
from sqlalchemy.orm import declarative_base, sessionmaker, Session

Base = declarative_base()


class Table_(Base):
    __tablename__ = project_name.replace('-', '_')

    date_created = Column(DateTime, default=None)
    date_edited = Column(DateTime, default=None)
    status = Column(String(16), default=None)
    payment_date = Column(DateTime, default=None)
    payment_id = Column(String(512), primary_key=True, default=None)
    payment_sum = Column(Float, default=None)
    contragent = Column(String(512), default=None)
    branch = Column(String(512), default=None)
    invoice_id = Column(Boolean, default=None)
    invoice_payment_to_contragent = Column(Boolean, default=None)
    tmz_realization = Column(Boolean, default=None)
    invoice_factura = Column(Boolean, default=None)
    subconto = Column(Boolean, default=None)

    @property
    def dict(self):
        m = self.__dict__.copy()
        return m


def add_to_db(session: Session, status_: str, payment_date_: datetime, payment_id_: str, payment_sum_: float, contragent_: str,
              branch_: str or None, invoice_id_: str or None, invoice_payment_to_contragent_: bool or None, tmz_realization_: bool or None,
              invoice_factura_: bool or None, subconto_: bool or None):
    session.add(Table_(
        date_created=datetime.datetime.now(),
        date_edited=None,
        status=status_,
        payment_date=payment_date_,
        payment_id=payment_id_,
        payment_sum=payment_sum_,
        contragent=contragent_,
        branch=branch_ if isinstance(branch_, str) else None,
        invoice_id=invoice_id_,
        invoice_payment_to_contragent=invoice_payment_to_contragent_,
        tmz_realization=tmz_realization_,
        invoice_factura=invoice_factura_,
        subconto=subconto_
    ))

    session.commit()


def get_all_data(session: Session):
    rows = [row for row in session.query(Table_).all()]

    print(type(rows[0]))

    return rows


def get_all_data_by_status(session: Session, status: str = 'new' or 'processing' or 'failed' or 'success'):
    rows = [row for row in session.query(Table_).filter(Table_.status == status).all()]

    return rows


def update_in_db(session: Session, row: Table_, status_: str, branch_: str or None, invoice_id_: str or None,
                 invoice_payment_to_contragent_: bool or None, tmz_realization_: bool or None, invoice_factura_: bool or None, subconto_: bool or None = None):
    row.status = status_
    row.date_edited = datetime.datetime.now()
    row.branch = branch_
    if invoice_id_ is not None:
        row.invoice_id = invoice_id_
    if invoice_payment_to_contragent_ is not None:
        row.invoice_payment_to_contragent = invoice_payment_to_contragent_
    if tmz_realization_ is not None:
        row.tmz_realization = tmz_realization_
    if invoice_factura_ is not None:
        row.invoice_factura = invoice_factura_
    if subconto_ is not None:
        row.subconto = subconto_

    session.commit()
