"""FORM 9 CONTROLLER - ang lohika ay nasa base_controller.py; dito ang mga setting ng Form 9 lang."""
import config_form9
from controllers.base_controller import BaseController
from services.form9_engine import Form9Engine


class Form9Controller(BaseController):
    engine_cls = Form9Engine
    page_title = config_form9.PAGE_TITLE
    template_name = config_form9.TEMPLATE_NAME
    template_keywords = config_form9.TEMPLATE_KEYWORDS
    output_folder_name = config_form9.OUTPUT_FOLDER_NAME
    template_label = "Form 9 Template"
    item_label = "Form 9"