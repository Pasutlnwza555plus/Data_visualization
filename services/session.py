from streamlit.runtime.state.session_state_proxy import SessionStateProxy
from enum import Enum

class SessionStateEnum(Enum):
    REFERENCE_SHEET = "reference_sheet"
    EOL_DATA = "eol_data"

class SessionStateManager:
    def __init__(self, session_state: SessionStateProxy):
        self._state: SessionStateProxy = session_state

    def _resolve_key(self, key):
        if not isinstance(key, SessionStateEnum):
            raise KeyError(f"state must be a SessionStateEnum.")
        return key.value

    def __getitem__(self, key):
        key_name = self._resolve_key(key)
        return self._state.get(key_name, None)

    def __setitem__(self, key, value):
        key_name = self._resolve_key(key)
        self._state[key_name] = value

    def get(self, key, default=None):
        return self._state.get(self._resolve_key(key), default)

    def keys(self):
        return list(self._state.keys())

    def items(self):
        return self._state.items()
