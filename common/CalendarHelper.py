__author__ = 'Max'
from  datetime import datetime, timedelta

class Period:
    def __init__(self,s,e):
        self.start=s
        self.end=e
def month_next():
    now=datetime.now()
    if now.month<12:
        s=datetime(now.year, now.month+1,1)
    else:
        s=datetime(now.year+1, 1,1)

    if s.month<12:
        e=datetime(s.year, s.month+1,1)
    else:
        e=datetime(s.year+1, 1,1)

    return Period(s,e-timedelta(seconds=1))
def month_current():
    m_t=datetime.now()
    m_cur_d_start=datetime(m_t.year,m_t.month,1,0,0,0)
    m_i=m_cur_d_start+timedelta(days=32)
    m_cur_d_finish=datetime(m_i.year,m_i.month,1,0,0,0)-timedelta(seconds=1)
    m_t=m_cur_d_start-timedelta(seconds=1)
    m_prev_d_finish=m_t
    m_prev_d_start=datetime(m_t.year,m_t.month,1)

    #self.month_prev= CalendarHelper.Period(m_prev_d_start,m_prev_d_finish)
    month_cur= Period(m_cur_d_start,m_cur_d_finish)
    return month_cur

def month_prev():
    m_t=datetime.now()
    m_cur_d_start=datetime(m_t.year,m_t.month,1,0,0,0)
    #m_i=m_cur_d_start+timedelta(days=32)
    #m_cur_d_finish=datetime(m_i.year,m_i.month,1,0,0,0)-timedelta(seconds=1)
    m_t=m_cur_d_start-timedelta(seconds=1)
    m_prev_d_finish=m_t
    m_prev_d_start=datetime(m_t.year,m_t.month,1)

    #self.month_prev= CalendarHelper.Period(m_prev_d_start,m_prev_d_finish)
    month_cur= Period(m_prev_d_start,m_prev_d_finish)
    return month_cur

class CalendarHelper22:

    def __init__(self):
        m_t=datetime.now()
        m_cur_d_start=datetime(m_t.year,m_t.month,1,0,0,0)
        m_i=m_cur_d_start+timedelta(days=32)
        m_cur_d_finish=datetime(m_i.year,m_i.month,1,0,0,0)-timedelta(seconds=1)
        m_t=m_cur_d_start-timedelta(seconds=1)
        m_prev_d_finish=m_t
        m_prev_d_start=datetime(m_t.year,m_t.month,1)

        self.month_prev= CalendarHelper.Period(m_prev_d_start,m_prev_d_finish)
        self.month_cur= CalendarHelper.Period(m_cur_d_start,m_cur_d_finish)