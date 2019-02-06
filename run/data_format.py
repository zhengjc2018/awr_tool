from app.commons import exceptions


class DataFormat:

    def success_(data):
        result = {
            "code": exceptions.OperationSuccess.business_code,
            "result": data,
            "message": "操作成功",
        }
        return result

    def failed_(data):
        result = {
            "code": exceptions.QueryFail.business_code,
            "result": [],
            "message": data,
        }
        return result
