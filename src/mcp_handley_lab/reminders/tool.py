"""Apple Reminders tool for managing reminders via MCP."""

import json
import os
import subprocess
from datetime import datetime
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field

from mcp_handley_lab.shared.models import OperationResult, ServerInfo

mcp = FastMCP("Apple Reminders Tool")

# Security: Restrict access to a single list if configured
ALLOWED_LIST = os.environ.get("REMINDERS_ALLOWED_LIST", "")
print(ALLOWED_LIST)


def _check_list_access(list_name: str) -> None:
    """Raise error if list access is not allowed."""
    if ALLOWED_LIST and list_name != ALLOWED_LIST:
        raise PermissionError(
            f"Access denied. Only the '{ALLOWED_LIST}' list is accessible. "
            f"Set REMINDERS_ALLOWED_LIST environment variable to change this."
        )


class Reminder(BaseModel):
    """A reminder item."""

    id: str = Field(..., description="Unique identifier for the reminder.")
    name: str = Field(..., description="Title/name of the reminder.")
    body: str = Field(default="", description="Notes/body text of the reminder.")
    completed: bool = Field(
        default=False, description="Whether the reminder is completed."
    )
    due_date: str = Field(
        default="", description="Due date in ISO format (YYYY-MM-DD HH:MM:SS)."
    )
    priority: int = Field(
        default=0, description="Priority: 0=none, 1=high, 5=medium, 9=low."
    )
    list_name: str = Field(
        default="", description="Name of the list containing this reminder."
    )


class ReminderList(BaseModel):
    """A reminders list."""

    id: str = Field(..., description="Unique identifier for the list.")
    name: str = Field(..., description="Name of the list.")
    reminder_count: int = Field(
        default=0, description="Number of reminders in the list."
    )


class ReminderSearchResult(BaseModel):
    """Search result for reminders."""

    query: str = Field(..., description="The search query executed.")
    reminders: list[Reminder] = Field(..., description="List of matching reminders.")
    total_found: int = Field(..., description="Total number of reminders found.")


def _run_jxa(script: str) -> Any:
    """Execute JXA script and return parsed JSON result."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"JXA error: {result.stderr.strip()}")
    output = result.stdout.strip()
    if not output:
        return None
    return json.loads(output)


@mcp.tool(
    description="Lists all reminder lists available in Apple Reminders. Returns list names and IDs."
)
def list_lists() -> list[ReminderList]:
    """Get all reminder lists."""
    script = """
    const app = Application('Reminders');
    const lists = app.lists();
    const result = lists.map(l => ({
        id: l.id(),
        name: l.name(),
        reminder_count: l.reminders().length
    }));
    JSON.stringify(result);
    """
    data = _run_jxa(script) or []
    lists = [ReminderList(**item) for item in data]
    if ALLOWED_LIST:
        lists = [l for l in lists if l.name == ALLOWED_LIST]
    return lists


@mcp.tool(
    description="Gets all reminders from a specific list, or all incomplete reminders if no list specified. Use `include_completed=True` to include completed reminders."
)
def get_reminders(
    list_name: str = Field(
        "",
        description="Name of the reminder list. If empty, returns reminders from all lists.",
    ),
    include_completed: bool = Field(
        False,
        description="Include completed reminders in results.",
    ),
) -> list[Reminder]:
    """Get reminders from a list."""
    # Apply security restriction
    effective_list = list_name or ALLOWED_LIST
    if list_name:
        _check_list_access(list_name)

    filter_completed = "false" if not include_completed else "true"
    list_filter = f'l.name() === "{effective_list}"' if effective_list else "true"

    script = f"""
    const app = Application('Reminders');
    const lists = app.lists().filter(l => {list_filter});
    const result = [];
    lists.forEach(list => {{
        list.reminders().forEach(r => {{
            if ({filter_completed} || !r.completed()) {{
                let dueDate = '';
                try {{ dueDate = r.dueDate() ? r.dueDate().toISOString() : ''; }} catch(e) {{}}
                result.push({{
                    id: r.id(),
                    name: r.name(),
                    body: r.body() || '',
                    completed: r.completed(),
                    due_date: dueDate,
                    priority: r.priority(),
                    list_name: list.name()
                }});
            }}
        }});
    }});
    JSON.stringify(result);
    """
    data = _run_jxa(script) or []
    return [Reminder(**item) for item in data]


@mcp.tool(
    description="Creates a new reminder in the specified list. Supports setting name, notes, due date, and priority."
)
def create_reminder(
    name: str = Field(..., description="Title of the reminder."),
    list_name: str = Field(
        "Reminders",
        description="Name of the list to add the reminder to. Defaults to 'Reminders'.",
    ),
    body: str = Field("", description="Notes/body text for the reminder."),
    due_date: str = Field(
        "",
        description="Due date in ISO format (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS).",
    ),
    priority: int = Field(
        0,
        description="Priority: 0=none, 1=high, 5=medium, 9=low.",
    ),
) -> Reminder:
    """Create a new reminder."""
    # Apply security restriction
    effective_list = ALLOWED_LIST or list_name
    _check_list_access(effective_list) if ALLOWED_LIST else None
    due_date_js = ""
    if due_date:
        due_date_js = f'props.dueDate = new Date("{due_date}");'

    script = f"""
    const app = Application('Reminders');
    const list = app.lists().find(l => l.name() === "{effective_list}");
    if (!list) throw new Error("List not found: {effective_list}");

    const props = {{
        name: "{name}",
        body: "{body}",
        priority: {priority}
    }};
    {due_date_js}

    const reminder = app.Reminder(props);
    list.reminders.push(reminder);

    let dueDate = '';
    try {{ dueDate = reminder.dueDate() ? reminder.dueDate().toISOString() : ''; }} catch(e) {{}}

    JSON.stringify({{
        id: reminder.id(),
        name: reminder.name(),
        body: reminder.body() || '',
        completed: reminder.completed(),
        due_date: dueDate,
        priority: reminder.priority(),
        list_name: "{effective_list}"
    }});
    """
    data = _run_jxa(script)
    return Reminder(**data)


@mcp.tool(
    description="Updates an existing reminder by ID. Any provided field will be updated."
)
def update_reminder(
    reminder_id: str = Field(..., description="ID of the reminder to update."),
    name: str = Field("", description="New title. Leave empty to keep current."),
    body: str = Field("", description="New notes. Leave empty to keep current."),
    due_date: str = Field(
        "",
        description="New due date in ISO format. Use 'clear' to remove due date.",
    ),
    priority: int = Field(
        -1,
        description="New priority (0=none, 1=high, 5=medium, 9=low). Use -1 to keep current.",
    ),
    completed: bool = Field(
        False,
        description="Set to true to mark as completed.",
    ),
) -> Reminder:
    """Update an existing reminder."""
    updates = []
    if name:
        updates.append(f'r.name = "{name}";')
    if body:
        updates.append(f'r.body = "{body}";')
    if due_date == "clear":
        updates.append("r.dueDate = null;")
    elif due_date:
        updates.append(f'r.dueDate = new Date("{due_date}");')
    if priority >= 0:
        updates.append(f"r.priority = {priority};")
    if completed:
        updates.append("r.completed = true;")

    updates_js = "\n".join(updates)

    script = f"""
    const app = Application('Reminders');
    let found = null;
    let listName = '';
    app.lists().some(list => {{
        const r = list.reminders().find(r => r.id() === "{reminder_id}");
        if (r) {{
            {updates_js}
            found = r;
            listName = list.name();
            return true;
        }}
        return false;
    }});

    if (!found) throw new Error("Reminder not found: {reminder_id}");

    let dueDate = '';
    try {{ dueDate = found.dueDate() ? found.dueDate().toISOString() : ''; }} catch(e) {{}}

    JSON.stringify({{
        id: found.id(),
        name: found.name(),
        body: found.body() || '',
        completed: found.completed(),
        due_date: dueDate,
        priority: found.priority(),
        list_name: listName
    }});
    """
    data = _run_jxa(script)
    return Reminder(**data)


@mcp.tool(description="Marks a reminder as completed by its ID.")
def complete_reminder(
    reminder_id: str = Field(..., description="ID of the reminder to complete."),
) -> OperationResult:
    """Mark a reminder as completed."""
    script = f"""
    const app = Application('Reminders');
    let found = false;
    app.lists().some(list => {{
        const r = list.reminders().find(r => r.id() === "{reminder_id}");
        if (r) {{
            r.completed = true;
            found = true;
            return true;
        }}
        return false;
    }});
    JSON.stringify({{found: found}});
    """
    data = _run_jxa(script)
    if not data.get("found"):
        raise RuntimeError(f"Reminder not found: {reminder_id}")
    return OperationResult(
        status="success", message=f"Reminder {reminder_id} marked as completed"
    )


@mcp.tool(description="Deletes a reminder by its ID.")
def delete_reminder(
    reminder_id: str = Field(..., description="ID of the reminder to delete."),
) -> OperationResult:
    """Delete a reminder."""
    script = f"""
    const app = Application('Reminders');
    let found = false;
    app.lists().some(list => {{
        const idx = list.reminders().findIndex(r => r.id() === "{reminder_id}");
        if (idx >= 0) {{
            list.reminders()[idx].delete();
            found = true;
            return true;
        }}
        return false;
    }});
    JSON.stringify({{found: found}});
    """
    data = _run_jxa(script)
    if not data.get("found"):
        raise RuntimeError(f"Reminder not found: {reminder_id}")
    return OperationResult(status="success", message=f"Reminder {reminder_id} deleted")


@mcp.tool(description="Searches reminders by name or body text across all lists.")
def search_reminders(
    query: str = Field(
        ..., description="Text to search for in reminder names and bodies."
    ),
    include_completed: bool = Field(False, description="Include completed reminders."),
) -> ReminderSearchResult:
    """Search reminders by text."""
    filter_completed = "false" if not include_completed else "true"
    query_lower = query.lower()
    list_filter = f'list.name() === "{ALLOWED_LIST}"' if ALLOWED_LIST else "true"

    script = f"""
    const app = Application('Reminders');
    const query = "{query_lower}";
    const result = [];
    app.lists().filter(list => {list_filter}).forEach(list => {{
        list.reminders().forEach(r => {{
            if ({filter_completed} || !r.completed()) {{
                const name = (r.name() || '').toLowerCase();
                const body = (r.body() || '').toLowerCase();
                if (name.includes(query) || body.includes(query)) {{
                    let dueDate = '';
                    try {{ dueDate = r.dueDate() ? r.dueDate().toISOString() : ''; }} catch(e) {{}}
                    result.push({{
                        id: r.id(),
                        name: r.name(),
                        body: r.body() || '',
                        completed: r.completed(),
                        due_date: dueDate,
                        priority: r.priority(),
                        list_name: list.name()
                    }});
                }}
            }}
        }});
    }});
    JSON.stringify(result);
    """
    data = _run_jxa(script) or []
    reminders = [Reminder(**item) for item in data]
    return ReminderSearchResult(
        query=query, reminders=reminders, total_found=len(reminders)
    )


@mcp.tool(description="Creates a new reminder list.")
def create_list(
    name: str = Field(..., description="Name for the new reminder list."),
) -> ReminderList:
    """Create a new reminder list."""
    if ALLOWED_LIST:
        raise PermissionError(
            f"Creating new lists is disabled. Only the '{ALLOWED_LIST}' list is accessible."
        )
    script = f"""
    const app = Application('Reminders');
    const newList = app.List({{name: "{name}"}});
    app.lists.push(newList);
    JSON.stringify({{
        id: newList.id(),
        name: newList.name(),
        reminder_count: 0
    }});
    """
    data = _run_jxa(script)
    return ReminderList(**data)


@mcp.tool(description="Checks Apple Reminders tool status and availability.")
def server_info() -> ServerInfo:
    """Get server status."""
    try:
        script = """
        const app = Application('Reminders');
        const lists = app.lists();
        JSON.stringify({count: lists.length});
        """
        data = _run_jxa(script)
        list_count = data.get("count", 0)
        status = "active"
        deps = {"reminders_app": f"{list_count} lists available"}
        if ALLOWED_LIST:
            deps["access_restriction"] = f"Limited to '{ALLOWED_LIST}' list only"
    except Exception as e:
        status = "error"
        deps = {"reminders_app": str(e)}

    return ServerInfo(
        name="Apple Reminders Tool",
        version="1.0.0",
        status=status,
        capabilities=[
            "list_lists",
            "get_reminders",
            "create_reminder",
            "update_reminder",
            "complete_reminder",
            "delete_reminder",
            "search_reminders",
            "create_list",
        ],
        dependencies=deps,
    )
