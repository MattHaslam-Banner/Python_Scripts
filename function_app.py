import azure.functions as func
import datetime
import json
import logging

app = func.FunctionApp()

@app.timer_trigger(schedule="0 0 0 1 * *", arg_name="myTimer", run_on_startup=False,
              use_monitor=False) 
def ExSWISchoolsOrderForm(myTimer: func.TimerRequest) -> None:
    
    if myTimer.past_due:
        logging.info('The timer is past due!')

    logging.info('GP Test Timer')

@app.route(route="ExSWISalesOrderForm", auth_level=func.AuthLevel.FUNCTION)
def ExSWISalesOrderForm(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    name = req.params.get('name')
    if not name:
        try:
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
             "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
             status_code=200
        )