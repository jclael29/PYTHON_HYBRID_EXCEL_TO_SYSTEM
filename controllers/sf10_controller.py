"""SF10 CONTROLLER - ang lohika ay nasa base_controller.py; dito ang mga setting ng SF10 lang."""
import config
from controllers.base_controller import BaseController
from services.sf10_engine import SF10Engine


class SF10Controller(BaseController):
    engine_cls = SF10Engine
    page_title = "Form Automation"
    template_name = config.TEMPLATE_NAME
    template_keywords = ("SHS", "TEMPLATE", "SF10")
    output_folder_name = config.OUTPUT_FOLDER_NAME
    template_label = "SF10 Template"
    item_label = "SF10"