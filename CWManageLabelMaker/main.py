from flask import Flask, request, abort
import functions
import threading
import queue
from time import sleep
import logging
from waitress import serve
import pathlib
import tomllib
import sys

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    # running as an executable (aka frozen)
    exe_dir = pathlib.Path(sys.executable).parent
else:
    # running live
    exe_dir = pathlib.Path(__file__).parent

config_path = pathlib.Path.cwd() / exe_dir / "config.toml"
with config_path.open(mode="rb") as fp:
    config = tomllib.load(fp)

print(f"Config File loaded: {config_path}")

debug = config["APP"]["DEBUG"]  # Switch between Debug mode and Production Mode


if debug == False:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s:%(levelname)s:%(message)s",
        filename="Label_Printer.log",
    )
    logging.getLogger().setLevel(logging.INFO)
else:
    print("Debug Enabled")
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s:%(levelname)s:%(message)s",
        filename="Label_Printer_Debug.log",
    )


job_queue = (
    queue.Queue()
)  # Used to hold label_job functions as the io for the printer needs to complete a job one at a time or it causes issues.


def worker():  # Define a Worker to handel jobs put into the "job_queue"
    while True:
        item = job_queue.get()
        logging.debug(f"Worker started work on {item}")
        item()
        job_queue.task_done()


# Turn-on the worker thread.
threading.Thread(target=worker, daemon=True).start()


app = Flask(__name__)


@app.route("/label", methods=["POST"])
def webhook():
    if request.method == "POST":
        request_json_data = request.json
        job_queue.put(lambda: functions.proccess_request(request_json_data))
        logging.info(
            f"Recieved ConnectWise Callback for {request_json_data['ID']} Sending to job queue, {job_queue.qsize()} tasks in queue."
        )
        logging.debug(f"Webhook JSON data:\n{request_json_data}")
        return (
            f"Task for PurchaseOrder:{request_json_data['ID']} added to queue. {job_queue.qsize()} tasks in queue.",
            200,
        )

    else:
        logging.info(f"Request: {request} was aborted")
        abort(400)


if __name__ == "__main__" and debug != True:
    logging.info("Started webserver in Production mode using 'Waitress'")
    serve(app, host="0.0.0.0", port=config["APP"]["PORT"])
else:
    logging.info("Started webserver in Developer mode using 'Flask Development Server'")
    app.run(use_reloader=False, debug=True)
