__author__ = 'Max'
from  datetime import datetime, timedelta
class CalendarHelper:
    class Period:
        def __init__(self,s,e):
            self.start=s
            self.end=e

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