"""
SAP GUI Scripting Framework
==============================
Generic Python wrapper built on the SAP GUI Scripting API.

This module provides reusable building blocks that map directly to the
SAP GUI Scripting API object model (GuiApplication → GuiConnection →
GuiSession → GuiMainWindow → components).  Every public helper is a
thin, documented layer over a documented COM method so you can compose
any transaction automation without hard-coding screen IDs.

API reference used:  SAP GUI Scripting API 6.40+

Requirements:
    pip install pywin32

Typical usage:
    from sap_scripting import SapSession
    with SapSession() as sap:
        sap.start_transaction("SE16H")
        sap.set_field("GD-TAB", "BSIK")
        sap.send_vkey(0)
        ...
"""

import win32com.client
import time
import logging
from typing import Optional, List, Dict, Any, Tuple
from contextlib import contextmanager

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("sap_scripting")


# ===========================================================================
#  CORE:  GuiApplication / GuiConnection / GuiSession wrapper
# ===========================================================================

class SapSession:
    """
    Wraps the SAP GUI Scripting runtime hierarchy:
        GuiApplication  →  GuiConnection  →  GuiSession

    API methods used:
        GetObject("SAPGUI")          – Running Object Table entry
        GetScriptingEngine           – returns GuiApplication
        Children(n)                  – navigate hierarchy
        Info  (GuiSessionInfo)       – session metadata
        StartTransaction / EndTransaction / SendCommand
        findById / findByName / findAllByName
        sendVKey
        LockSessionUI / UnlockSessionUI
    """

    # VKey constants (from SAP documentation)
    VKEY_ENTER   = 0
    VKEY_F2      = 2
    VKEY_F3      = 3   # Back
    VKEY_F5      = 5
    VKEY_F6      = 6
    VKEY_F7      = 7
    VKEY_F8      = 8   # Execute
    VKEY_F9      = 9
    VKEY_F12     = 12  # Cancel
    VKEY_SHIFT_F4 = 16  # Save As
    VKEY_CTRL_SHIFT_F12 = 36

    def __init__(self, connection_index: int = 0, session_index: int = 0):
        self.connection_index = connection_index
        self.session_index = session_index
        self.application = None  # GuiApplication
        self.connection = None   # GuiConnection
        self.session = None      # GuiSession
        self._connect()

    # -- Context manager support --
    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass  # Session stays open; SAP manages lifetime

    # ------------------------------------------------------------------
    #  Connection  (GuiApplication → GuiConnection → GuiSession)
    # ------------------------------------------------------------------

    def _connect(self):
        """
        Attach to a running SAP GUI via the Running Object Table.

        Uses:
            GetObject("SAPGUI")      → ROT entry
            .GetScriptingEngine      → GuiApplication
            .Children(n)             → GuiConnection / GuiSession
        """
        try:
            rot_entry = win32com.client.GetObject("SAPGUI")
            self.application = rot_entry.GetScriptingEngine
            self.connection = self.application.Children(self.connection_index)
            self.session = self.connection.Children(self.session_index)
            info = self.session.Info
            log.info(
                f"Connected: system={info.SystemName}, "
                f"client={info.Client}, user={info.User}, "
                f"transaction={info.Transaction}"
            )
        except Exception as e:
            raise ConnectionError(
                f"Cannot connect to SAP GUI: {e}\n"
                "  • Is SAP GUI running and logged in?\n"
                "  • Is scripting enabled (Alt+F12 → Options → Scripting)?\n"
                "  • Is sapgui/user_scripting = TRUE on the server?"
            )

    def get_session_info(self) -> Dict[str, str]:
        """
        Read GuiSessionInfo properties.

        Uses: session.Info.* (SystemName, Client, User, Language,
              Transaction, Program, ScreenNumber, ApplicationServer, etc.)
        """
        info = self.session.Info
        return {
            "system": info.SystemName,
            "client": info.Client,
            "user": info.User,
            "language": info.Language,
            "transaction": info.Transaction,
            "program": info.Program,
            "screen": str(info.ScreenNumber),
            "response_time": str(info.ResponseTime),
            "round_trips": str(info.RoundTrips),
        }

    # ------------------------------------------------------------------
    #  Navigation  (SendCommand, StartTransaction, EndTransaction)
    # ------------------------------------------------------------------

    def start_transaction(self, tcode: str):
        """
        Navigate to a transaction.

        Uses: GuiSession.StartTransaction(tcode)
              Equivalent to SendCommand("/n<tcode>")
        """
        log.info(f"Starting transaction: {tcode}")
        self.session.StartTransaction(tcode)

    def end_transaction(self):
        """
        End current transaction (returns to menu).

        Uses: GuiSession.EndTransaction()
              Equivalent to SendCommand("/n")
        """
        self.session.EndTransaction()

    def send_command(self, command: str):
        """
        Execute any command string (same as typing in the OKCode field).

        Uses: GuiSession.SendCommand(command)
        Examples: "/nSE16", "/nend", "/nex", "/o" (new session)
        """
        log.info(f"SendCommand: {command}")
        self.session.SendCommand(command)

    def send_vkey(self, vkey: int, window_id: str = "wnd[0]"):
        """
        Emulate pressing a virtual key (Enter, F8, etc.).

        Uses: GuiFrameWindow.sendVKey(vkey)
        Common keys: 0=Enter, 2=F2, 3=Back, 8=F8/Execute, 12=Cancel
        """
        self.session.findById(window_id).sendVKey(vkey)

    # ------------------------------------------------------------------
    #  Element access  (findById, findByName, findAllByName, Children)
    # ------------------------------------------------------------------

    def find_by_id(self, element_id: str, raise_error: bool = True):
        """
        Locate a UI element by its full scripting ID.

        Uses: GuiSession.findById(id)
              The id is a URL-like path: "wnd[0]/usr/ctxtFIELD_NAME"

        Returns the COM object, or None if raise_error=False and not found.
        """
        try:
            return self.session.findById(element_id)
        except Exception:
            if raise_error:
                raise
            return None

    def find_by_name(self, name: str, component_type: str = ""):
        """
        Find the first element matching a SAP data dictionary field name.

        Uses: GuiComponent.findByName(name, type)
              type is the type prefix: "txt", "ctxt", "chk", "rad", "btn", etc.

        This is useful when you know the field name but not the full path.
        """
        if component_type:
            return self.session.findByName(name, component_type)
        return self.session.findByName(name, "")

    def find_all_by_name(self, name: str, component_type: str = "") -> list:
        """
        Find ALL elements matching a name (returns a GuiComponentCollection).

        Uses: GuiComponent.findAllByName(name, type)
              Unlike findByName which returns only the first match.
        """
        collection = self.session.findAllByName(name, component_type)
        return [collection(i) for i in range(collection.Count)]

    def get_children(self, element_id: str = "wnd[0]/usr") -> list:
        """
        Get all direct children of a container element.

        Uses: GuiVContainer.Children  (GuiComponentCollection)
        Default scans the user area of the main window.
        """
        container = self.session.findById(element_id)
        children = container.Children
        return [children(i) for i in range(children.Count)]

    # ------------------------------------------------------------------
    #  Field interaction  (Text, Selected, Press, Select)
    # ------------------------------------------------------------------

    def set_field(self, field_name: str, value: str, field_type: str = ""):
        """
        Set a text/input field value by SAP data dictionary name.

        Uses: GuiTextField.text = value  (or GuiCTextField.text)
              Finds the field via findByName, then sets .text property.
        """
        if field_type:
            element = self.session.findByName(field_name, field_type)
        else:
            # Try common field types
            for ftype in ["ctxt", "txt", "cmb"]:
                try:
                    element = self.session.findByName(field_name, ftype)
                    break
                except Exception:
                    continue
            else:
                raise ValueError(f"Field '{field_name}' not found as ctxt/txt/cmb")

        element.text = value
        log.debug(f"Set field {field_name} = '{value}'")

    def set_field_by_id(self, element_id: str, value: str):
        """
        Set a field value by its full scripting ID.

        Uses: findById(id).text = value
        """
        self.session.findById(element_id).text = value

    def get_field(self, field_name: str, field_type: str = "") -> str:
        """
        Read a field's current text value by name.

        Uses: GuiTextField.text  (read-only for GuiPasswordField)
        """
        if field_type:
            return self.session.findByName(field_name, field_type).text
        for ftype in ["txt", "ctxt", "lbl"]:
            try:
                return self.session.findByName(field_name, ftype).text
            except Exception:
                continue
        raise ValueError(f"Field '{field_name}' not found")

    def get_field_by_id(self, element_id: str) -> str:
        """Read a field value by its full scripting ID."""
        return self.session.findById(element_id).text

    def set_checkbox(self, field_name: str, checked: bool = True):
        """
        Set a checkbox state.

        Uses: GuiCheckBox.selected = True/False
              Type prefix: chk
        """
        cb = self.session.findByName(field_name, "chk")
        cb.selected = checked

    def set_checkbox_by_id(self, element_id: str, checked: bool = True):
        """Set a checkbox state by full ID."""
        self.session.findById(element_id).selected = checked

    def select_radio(self, field_name: str):
        """
        Select a radio button.

        Uses: GuiRadioButton.select()
              Type prefix: rad
        """
        self.session.findByName(field_name, "rad").select()

    def press_button(self, field_name: str = "", element_id: str = ""):
        """
        Press a button.

        Uses: GuiButton.press()
              Type prefix: btn
        """
        if element_id:
            self.session.findById(element_id).press()
        elif field_name:
            self.session.findByName(field_name, "btn").press()

    def select_tab(self, tab_id: str):
        """
        Select a tab on a tab strip.

        Uses: GuiTab.select()
              Type prefix: tabp
        """
        self.session.findById(tab_id).select()

    def select_combo_entry(self, field_name: str, key: str):
        """
        Select an entry in a combo box by key.

        Uses: GuiComboBox.key = value
              GuiComboBox.Entries contains GuiComboBoxEntry items
        """
        combo = self.session.findByName(field_name, "cmb")
        combo.key = key

    # ------------------------------------------------------------------
    #  Status bar  (GuiStatusbar)
    # ------------------------------------------------------------------

    def get_statusbar(self) -> Dict[str, str]:
        """
        Read the status bar message.

        Uses: GuiStatusbar.Text, .MessageType
              MessageType: S=Success, W=Warning, E=Error, A=Abort, I=Info
        """
        sbar = self.session.findById("wnd[0]/sbar")
        return {
            "text": sbar.Text,
            "type": sbar.MessageType,
        }

    def check_statusbar_error(self) -> Optional[str]:
        """Returns error message text if statusbar shows an error, else None."""
        sb = self.get_statusbar()
        if sb["type"] in ("E", "A"):
            return sb["text"]
        return None

    # ------------------------------------------------------------------
    #  Modal window / pop-up handling  (GuiModalWindow)
    # ------------------------------------------------------------------

    def handle_popup(self, button_id: str = "wnd[1]/tbar[0]/btn[0]",
                     check_exists: bool = True) -> bool:
        """
        Try to dismiss a modal popup window.

        Uses: GuiModalWindow (wnd[1]) — press a button on it.
        Returns True if a popup was handled, False if none existed.
        """
        if check_exists:
            popup = self.find_by_id("wnd[1]", raise_error=False)
            if popup is None:
                return False
        try:
            self.session.findById(button_id).press()
            return True
        except Exception:
            return False

    # ------------------------------------------------------------------
    #  Grid / ALV  (GuiGridView)
    # ------------------------------------------------------------------

    def find_grid(self, search_id: str = "wnd[0]/usr"):
        """
        Search for a GuiGridView control within a container.

        Uses: Children traversal, checking .type == "GuiGridView"
              Falls back to common known grid paths.

        Returns the grid COM object or None.
        """
        # Try common grid container paths first
        common_paths = [
            f"{search_id}/cntlRESULT_LIST/shellcont/shell",
            f"{search_id}/cntlCONTAINER/shellcont/shell",
            f"{search_id}/cntlGRID1/shellcont/shell",
            f"{search_id}/cntlGRID/shellcont/shell",
            f"{search_id}/cntlALV_CONTAINER/shellcont/shell",
        ]
        for path in common_paths:
            grid = self.find_by_id(path, raise_error=False)
            if grid is not None:
                return grid
        return None

    def grid_get_row_count(self, grid) -> int:
        """
        Uses: GuiGridView.RowCount
        """
        return grid.RowCount

    def grid_get_column_count(self, grid) -> int:
        """
        Uses: GuiGridView.ColumnCount
        """
        return grid.ColumnCount

    def grid_get_columns(self, grid) -> List[str]:
        """
        Get all column identifiers in display order.

        Uses: GuiGridView.ColumnOrder (collection of column ID strings)
        """
        col_order = grid.ColumnOrder
        return [col_order(i) for i in range(col_order.Count)]

    def grid_get_column_titles(self, grid) -> Dict[str, str]:
        """
        Map column IDs to their display titles.

        Uses: GuiGridView.GetColumnTitles(col_id) — returns collection
              GuiGridView.GetDisplayedColumnTitle(col_id)
        """
        columns = self.grid_get_columns(grid)
        titles = {}
        for col in columns:
            try:
                titles[col] = grid.GetDisplayedColumnTitle(col)
            except Exception:
                titles[col] = col
        return titles

    def grid_get_cell_value(self, grid, row: int, column: str) -> str:
        """
        Uses: GuiGridView.GetCellValue(row, column)
        """
        return grid.GetCellValue(row, column)

    def grid_read_all(self, grid, columns: List[str] = None,
                      max_rows: int = 0) -> List[Dict[str, str]]:
        """
        Read all data from a GuiGridView into a list of dicts.

        Uses: RowCount, ColumnOrder, GetCellValue

        Args:
            grid:     The GuiGridView COM object
            columns:  Specific columns to read (default: all)
            max_rows: Limit rows (0 = all)
        """
        if columns is None:
            columns = self.grid_get_columns(grid)

        total = grid.RowCount
        if max_rows > 0:
            total = min(total, max_rows)

        data = []
        for row in range(total):
            record = {}
            for col in columns:
                try:
                    record[col] = grid.GetCellValue(row, col)
                except Exception:
                    record[col] = ""
            data.append(record)

        log.info(f"Grid read: {total} rows × {len(columns)} columns")
        return data

    def grid_get_distinct(self, grid, column: str) -> List[str]:
        """
        Get distinct values for a single column from a grid.

        Uses: RowCount, GetCellValue
        """
        values = set()
        for row in range(grid.RowCount):
            val = grid.GetCellValue(row, column).strip()
            values.add(val)
        return sorted(values)

    def grid_click_cell(self, grid, row: int, column: str):
        """
        Uses: GuiGridView.Click(row, column)
        """
        grid.Click(row, column)

    def grid_double_click_cell(self, grid, row: int, column: str):
        """
        Uses: GuiGridView.DoubleClick(row, column)
        """
        grid.DoubleClick(row, column)

    def grid_select_rows(self, grid, rows: str):
        """
        Select rows in grid.

        Uses: GuiGridView.SelectedRows = "0,1,3-5"
        """
        grid.SelectedRows = rows

    def grid_set_current_cell(self, grid, row: int, column: str):
        """
        Uses: GuiGridView.CurrentCellRow, .CurrentCellColumn, .CurrentCellMoved
        """
        grid.CurrentCellRow = row
        grid.CurrentCellColumn = column

    # ------------------------------------------------------------------
    #  Table control  (GuiTableControl — different from GuiGridView)
    # ------------------------------------------------------------------

    def table_get_row_count(self, table) -> int:
        """
        Uses: GuiTableControl.RowCount
        """
        return table.RowCount

    def table_get_columns(self, table) -> list:
        """
        Uses: GuiTableControl.Columns  (GuiTableColumn collection)
        """
        cols = table.Columns
        return [cols(i) for i in range(cols.Count)]

    def table_read_cell(self, table_id: str, row: int, col_name: str) -> str:
        """
        Read a cell from a dynpro table control.

        Table controls use a different structure than GridView:
        The cell path is: <table_id>/ctxt<FIELDNAME>[col_idx, row_idx]
        """
        try:
            cell = self.session.findById(f"{table_id}/ctxt{col_name}[0,{row}]")
            return cell.text
        except Exception:
            try:
                cell = self.session.findById(f"{table_id}/txt{col_name}[0,{row}]")
                return cell.text
            except Exception:
                return ""

    # ------------------------------------------------------------------
    #  Screen exploration  (DumpState, Visualize, Type introspection)
    # ------------------------------------------------------------------

    def explore_screen(self, container_id: str = "wnd[0]/usr") -> List[Dict]:
        """
        Walk all children of a container and collect their properties.
        Useful for discovering field IDs and types on unknown screens.

        Uses: Children collection, .Id, .Type, .Name, .Text (where available)
        """
        elements = []
        try:
            container = self.session.findById(container_id)
            children = container.Children
            for i in range(children.Count):
                child = children(i)
                info = {
                    "id": child.Id,
                    "type": child.Type,
                    "name": child.Name,
                }
                try:
                    info["text"] = child.Text
                except Exception:
                    info["text"] = ""
                try:
                    info["changeable"] = child.Changeable
                except Exception:
                    info["changeable"] = None

                elements.append(info)

                # Recurse one level for containers
                try:
                    if child.ContainerType:
                        sub_children = child.Children
                        for j in range(sub_children.Count):
                            sc = sub_children(j)
                            sub_info = {
                                "id": sc.Id,
                                "type": sc.Type,
                                "name": sc.Name,
                            }
                            try:
                                sub_info["text"] = sc.Text
                            except Exception:
                                sub_info["text"] = ""
                            elements.append(sub_info)
                except Exception:
                    pass

        except Exception as e:
            log.warning(f"explore_screen error: {e}")

        return elements

    def visualize_element(self, element_id: str, on: bool = True):
        """
        Draw a red frame around a UI element (for debugging).

        Uses: GuiVComponent.Visualize(on)
        """
        self.session.findById(element_id).Visualize(on)

    # ------------------------------------------------------------------
    #  Session locking
    # ------------------------------------------------------------------

    def lock_ui(self):
        """
        Lock session to prevent user interaction during script execution.

        Uses: GuiSession.LockSessionUI()
        """
        self.session.LockSessionUI()

    def unlock_ui(self):
        """
        Unlock session after script execution.

        Uses: GuiSession.UnlockSessionUI()
        """
        self.session.UnlockSessionUI()

    @contextmanager
    def locked(self):
        """Context manager: locks UI during block, always unlocks after."""
        self.lock_ui()
        try:
            yield
        finally:
            self.unlock_ui()


# ===========================================================================
#  HELPERS:  Generic transaction patterns
# ===========================================================================

def run_se16h(sap: SapSession, table: str,
              fields: List[Dict[str, str]],
              max_rows: int = 0) -> List[Dict[str, str]]:
    """
    Generic SE16H execution.

    Args:
        sap:    SapSession instance
        table:  SAP table name (e.g. "BSIK", "BSID", "BSIS", "EKKO")
        fields: List of field configs, each dict:
                  { "name": "UMSKZ",       ← SAP field name
                    "group": True,          ← group/aggregate (optional)
                    "sum": True,            ← sum (optional)
                    "value": "A" }          ← selection value (optional)
        max_rows: Limit rows returned (0 = all)

    Returns:
        List of row dicts with field values.
    """
    sap.start_transaction("SE16H")
    sap.set_field_by_id("wnd[0]/usr/ctxtGD-TAB", table)
    sap.send_vkey(SapSession.VKEY_ENTER)

    # Fill fields into the selection table
    for idx, field_cfg in enumerate(fields):
        fname = field_cfg["name"]
        # Set field name in the fields table
        sap.set_field_by_id(
            f"wnd[0]/usr/tblSAPLSE16HFIELDS_TABLE/ctxtGS_FIELDS-FIELDNAME[0,{idx}]",
            fname,
        )
        # Group checkbox (column index 4 in SE16H fields table)
        if field_cfg.get("group"):
            sap.set_checkbox_by_id(
                f"wnd[0]/usr/tblSAPLSE16HFIELDS_TABLE/chkGS_FIELDS-AGGR[4,{idx}]",
                True,
            )
        # Sum checkbox (column index 5 in SE16H fields table)
        if field_cfg.get("sum"):
            sap.set_checkbox_by_id(
                f"wnd[0]/usr/tblSAPLSE16HFIELDS_TABLE/chkGS_FIELDS-SUM[5,{idx}]",
                True,
            )

    sap.send_vkey(SapSession.VKEY_ENTER)

    # Set selection values (if any fields have a "value" key)
    for field_cfg in fields:
        if "value" in field_cfg:
            try:
                sap.set_field(field_cfg["name"], field_cfg["value"])
            except Exception:
                pass

    # Set max rows if specified
    if max_rows > 0:
        try:
            sap.set_field_by_id("wnd[0]/usr/txtGD-MAX_LINES", str(max_rows))
        except Exception:
            pass

    # Execute
    sap.send_vkey(SapSession.VKEY_F8)

    # Handle possible popups (e.g. "maximum number of entries")
    sap.handle_popup()

    # Check for errors
    err = sap.check_statusbar_error()
    if err:
        log.warning(f"SE16H status: {err}")
        return []

    # Read grid
    grid = sap.find_grid()
    if grid is None:
        log.warning("SE16H: No grid found in results")
        return []

    # Determine which columns to read
    read_cols = [f["name"] for f in fields]
    # Also try to read COUNT column if grouping was used
    if any(f.get("group") for f in fields):
        read_cols.append("COUNT")

    return sap.grid_read_all(grid, columns=read_cols, max_rows=max_rows)


def run_transaction_report(sap: SapSession, tcode: str,
                           selection_fields: Dict[str, str] = None,
                           execute_vkey: int = SapSession.VKEY_F8,
                           read_columns: List[str] = None,
                           max_rows: int = 0) -> List[Dict[str, str]]:
    """
    Generic: open a transaction, fill selection fields, execute, read grid.

    Works for report transactions like FBL1N, FBL3N, FBL5N, ME2M, VA05, etc.

    Args:
        sap:              SapSession instance
        tcode:            Transaction code
        selection_fields: Dict of {field_name: value} for selection screen
        execute_vkey:     VKey to execute (default F8)
        read_columns:     Columns to extract (None = all)
        max_rows:         Max rows to return (0 = all)
    """
    sap.start_transaction(tcode)

    # Fill selection screen
    if selection_fields:
        for field_name, value in selection_fields.items():
            try:
                sap.set_field(field_name, value)
            except Exception:
                try:
                    sap.set_field(field_name, value, "txt")
                except Exception:
                    log.warning(f"Could not set field {field_name}")

    # Execute
    sap.send_vkey(execute_vkey)

    # Handle popups
    sap.handle_popup()

    # Check status
    err = sap.check_statusbar_error()
    if err:
        log.warning(f"{tcode} status: {err}")
        return []

    # Read grid
    grid = sap.find_grid()
    if grid is None:
        log.warning(f"{tcode}: No grid found")
        return []

    return sap.grid_read_all(grid, columns=read_columns, max_rows=max_rows)
