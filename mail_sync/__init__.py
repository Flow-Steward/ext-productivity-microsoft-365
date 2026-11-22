"""New-mail trigger poll for Outlook connections (task B5.2).

- :mod:`mail_sync.state` holds the versioned sync state and time helpers.
- :mod:`mail_sync.machine` is the pure state machine: which request comes next,
  how a page changes the state, which messages become events, and recovery.
- :mod:`mail_sync.graph` wraps the Graph calls the machine asks for.
- :mod:`mail_sync.poll` is the ``poll_mailbox`` background action: it reads the
  connection and its state, runs one budgeted poll step and answers with the new
  state (a ``project_connection_upsert`` host mutation) and the events.

``docs/TRIGGER_CONTRACT.md`` describes exactly what goes to Core.
"""
