"""outlook_access — read emails from the Outlook desktop COM Automation Client.

Read-only by design: nothing in this package sends, drafts, deletes, or marks
messages as read. The surface is deliberately narrow so that later
functionality (labelling, triage, forwarding) can be added as separate modules
without this layer ever gaining write capabilities by accident.
"""

__all__ = ["com", "reader", "state", "doctor"]
