from app.models import AwrHistory, AwrResult, Hosts, AwrStateCode
from app.extensions import db


class DBfunc:
    def commit_func(obj):
        try:
            db.session.add(obj)
            db.session.commit()
            db.session.refresh(obj)
            return obj
        except Exception:
            db.session.rollback()

    def awr_history_insert(host_id, result_id, name, start_snap, end_snap):
        history = AwrHistory().insert(host_id, result_id, name, start_snap, end_snap)
        history = DBfunc.commit_func(history)

        if history:
            return history.id

    def awr_history_update(history_id, code):
        history = AwrHistory.query.get(history_id)
        history = history.update(code)
        history = DBfunc.commit_func(history)

        if history:
            return history.id

    def awr_result_insert(host_id, snap_start_at, snap_end_at):
        result = AwrResult().insert(host_id, snap_start_at, snap_end_at)
        result = DBfunc.commit_func(result)

        if result:
            return result.id

    def awr_result_update(result_id, history_id, code, snap_start, snap_end):
        result = AwrResult.query.get(result_id)
        result = result.update(history_id, code, snap_start, snap_end)
        result = DBfunc.commit_func(result)

        if result:
            return result.id

    def update_for_processing(host_id, start_time, end_time):
        host = Hosts.query.get(host_id)

        obj = host.awr_father
        result_id = obj[0].id if obj else \
            DBfunc.awr_result_insert(host_id, start_time, end_time)

        code = AwrStateCode.PROCESSING
        history_id = None
        DBfunc.awr_result_update(result_id, history_id, code, '', '')

        return result_id

    def update_for_finish(result_id, history_id, snap_start, snap_end):
        code = AwrStateCode.FINISH
        DBfunc.awr_result_update(result_id, history_id, code, snap_start, snap_end)

    def update_for_fail(result_id):
        code = AwrStateCode.FAIL
        history_id = None
        DBfunc.awr_result_update(result_id, history_id, code, '', '')

    def update_for_analyze_finish(history_id):
        code = AwrStateCode.FINISH
        DBfunc.awr_history_update(history_id, code)

    def update_for_analyze_fail(history_id):
        code = AwrStateCode.FAIL
        DBfunc.awr_history_update(history_id, code)

    def update_for_analyze_processing(history_id):
        code = AwrStateCode.PROCESSING
        DBfunc.awr_history_update(history_id, code)
