from enum import Enum, auto

class ChatState(Enum):
    AWAITING_COMMAND = auto()
    AWAITING_USER_INFO_NAME = auto()
    AWAITING_USER_INFO_EXTENSION = auto()
    AWAITING_USER_INFO_EMPLOYEE_ID = auto()
    AWAITING_DEVICE_TYPE = auto()
    AWAITING_DATES = auto()
    CONFIRM_RESERVATION = auto()
    COMPLETED = auto()
    CANCEL_CONFIRM = auto()
