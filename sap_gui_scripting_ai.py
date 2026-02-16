# SAP GUI AI Agent - Using SAP AI Core (Generative AI Hub) instead of OpenAI
#
# Requirements:
# pip install pywin32 langchain langgraph generative-ai-hub-sdk[langchain]

import pythoncom
import win32com.client
import typing
from typing import Annotated

from langchain_core.messages import AnyMessage
from langgraph.graph.message import add_messages
from langgraph.graph import END, StateGraph
from langgraph.prebuilt import ToolNode, tools_condition
from dotenv import load_dotenv
load_dotenv()  # Add this at the top
# --- SAP AI Core / Generative AI Hub Integration ---
from gen_ai_hub.proxy.langchain.init_models import init_llm

# Initialize LLM from SAP AI Core (choose any deployed model)
# Common options: "gpt-4o", "gpt-4o-mini", "claude-3-5-sonnet", "gemini-1.5-pro", etc.
# Check your AI Launchpad for exact model names and deployment IDs
llm = init_llm(
    model_name="gpt-4o",          # or any model you have deployed
    temperature=0.0,
    max_tokens=1024,
    # deployment_id="your-deployment-id"  # optional, if using custom deployment
)

# Bind tools later after defining them
# model_with_tools = llm.bind_tools(tools) will be done after tools are defined

# --- SAP GUI Automation Class ---
class SAPAutomation:
    def __init__(self):
        pythoncom.CoInitialize()
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")
            self.engine = self.sap_gui.GetScriptingEngine
            if self.engine.Children.Count == 0:
                raise ValueError("No SAP connections open.")
            self.connection = self.engine.Children(0)
            if self.connection.Children.Count == 0:
                raise ValueError("No SAP sessions open.")
            self.session = self.connection.Children(0)
        except Exception as e:
            raise RuntimeError(f"Failed to connect to SAP GUI: {str(e)}")

    def find_by_id(self, element_id: str):
        try:
            return self.session.FindById(element_id)
        except Exception:
            raise ValueError(f"Element not found: {element_id}")

    def start_transaction(self, transaction_code: str):
        self.session.StartTransaction(transaction_code)

    def set_text(self, element_id: str, value: str):
        element = self.find_by_id(element_id)
        element.Text = value

    def get_text(self, element_id: str) -> str:
        element = self.find_by_id(element_id)
        return element.Text

    def press_button(self, element_id: str):
        element = self.find_by_id(element_id)
        element.Press()

    def send_vkey(self, vkey: int):
        self.session.SendVKey(vkey)  # e.g., 0 for Enter

    def get_gui_state(self) -> str:
        window = self.session.ActiveWindow
        state = f"Current Window: {window.Text}\nElements:\n"
        def recurse_children(parent, indent=""):
            for i in range(parent.Children.Count):
                child = parent.Children(i)
                text = getattr(child, 'Text', 'N/A')
                state_line = f"{indent}- ID: {child.Id} | Type: {child.Type} | Text: {text}\n"
                state += state_line
                if child.Children.Count > 0:
                    recurse_children(child, indent + "  ")
        recurse_children(window)
        return state

# Instantiate globally (or pass via state in production)
sap_automation = SAPAutomation()

# --- Tools ---
from langchain.tools import tool

@tool
def start_transaction(transaction_code: str) -> str:
    """Start a new SAP transaction by code (e.g., 'MM01')."""
    sap_automation.start_transaction(transaction_code)
    return f"Transaction {transaction_code} started."

@tool
def set_field_value(element_id: str, value: str) -> str:
    """Set the text value of a field by its full ID."""
    sap_automation.set_text(element_id, value)
    return "Field value set."

@tool
def get_field_value(element_id: str) -> str:
    """Get the text value of a field by its full ID."""
    return sap_automation.get_text(element_id)

@tool
def press_button(element_id: str) -> str:
    """Press a button by its full ID."""
    sap_automation.press_button(element_id)
    return "Button pressed."

@tool
def send_enter() -> str:
    """Send the Enter key (VKey 0)."""
    sap_automation.send_vkey(0)
    return "Enter key sent."

@tool
def get_current_gui_state() -> str:
    """Return a textual description of all visible elements in the current window."""
    return sap_automation.get_gui_state()

# Collect tools
tools = [
    start_transaction,
    set_field_value,
    get_field_value,
    press_button,
    send_enter,
    get_current_gui_state
]

# Bind tools to LLM
model_with_tools = llm.bind_tools(tools)

# --- LangGraph State ---
class AgentState(typing.TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]

# --- Nodes ---
def reasoner_node(state: AgentState):
    messages = state["messages"]
    # Optionally inject current GUI state for better context
    gui_state = get_current_gui_state.invoke({})
    enhanced_messages = messages + [
        {"role": "system", "content": f"Current SAP GUI State:\n{gui_state}"}
    ]
    response = model_with_tools.invoke(enhanced_messages)
    return {"messages": [response]}

tool_executor = ToolNode(tools)

def tools_node(state: AgentState):
    last_message = state["messages"][-1]
    if not hasattr(last_message, "tool_calls") or not last_message.tool_calls:
        return {"messages": []}
    
    outputs = tool_executor.batch([
        {
            "tool_name": tc["name"],
            "tool_args": tc["args"],
            "tool_call_id": tc["id"]
        }
        for tc in last_message.tool_calls
    ])
    
    tool_messages = [
        {"role": "tool", "content": str(output), "tool_call_id": tc["id"]}
        for tc, output in zip(last_message.tool_calls, outputs)
    ]
    return {"messages": tool_messages}

# --- Build Graph ---
graph = StateGraph(AgentState)
graph.add_node("reasoner", reasoner_node)
graph.add_node("tools", tools_node)
graph.set_entry_point("reasoner")
graph.add_conditional_edges("reasoner", tools_condition)
graph.add_edge("tools", "reasoner")
graph.add_edge("reasoner", END)  # Optional: direct end if no tools

app = graph.compile()

# --- Run Agent ---
def run_agent(query: str):
    initial_state = {"messages": [{"role": "user", "content": query}]}
    result = None
    for step in app.stream(initial_state):
        result = step
    # Final answer is the last non-tool message
    final_messages = result.get("reasoner", {}).get("messages", [])
    if final_messages:
        return final_messages[-1].content
    return "Task completed."

# Example
if __name__ == "__main__":
    query = "Go to transaction SE38, create a new report named ZTEST_AGENT, add a simple WRITE statement, and activate it."
    response = run_agent(query)
    print("Final Response:", response)