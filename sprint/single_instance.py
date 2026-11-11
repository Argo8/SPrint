"""One SPrint window at a time: starting it again (e.g. from "Send to") hands
the files to the window that is already open."""

from __future__ import annotations

import json
import logging
import os
from typing import Callable

from PySide6.QtNetwork import QLocalServer, QLocalSocket

log = logging.getLogger(__name__)

CONNECT_MS = 1000
HANDOVER_MS = 3000   # a busy window may take a moment to read


class SingleInstance:
    def __init__(self, name: str):
        self.name = name
        self.server = None

    def send(self, paths: list[str]) -> bool:
        """Pass ``paths`` to a running SPrint; False if there is none (or it did not take them)."""
        socket = QLocalSocket()
        socket.connectToServer(self.name)
        if not socket.waitForConnected(CONNECT_MS):
            return False
        if os.name == "nt":   # let the running window come to the front
            try:
                import ctypes
                ctypes.windll.user32.AllowSetForegroundWindow(-1)
            except Exception:
                pass
        socket.write(json.dumps([os.path.abspath(p) for p in paths]).encode("utf-8") + b"\n")
        if not socket.waitForBytesWritten(HANDOVER_MS):   # not taken: open a window of our own
            socket.abort()
            return False
        # The window closes the connection once it has the paths; leaving earlier could lose them.
        if socket.state() == QLocalSocket.LocalSocketState.ConnectedState:
            socket.waitForDisconnected(HANDOVER_MS)
        socket.disconnectFromServer()
        return True

    def listen(self, on_paths: Callable[[list[str]], None]) -> None:
        self.server = QLocalServer()
        self.server.setSocketOptions(QLocalServer.SocketOption.UserAccessOption)
        if not self.server.listen(self.name):
            QLocalServer.removeServer(self.name)   # left over from a crash (not on Windows)
            if not self.server.listen(self.name):
                log.warning("Cannot listen for other SPrint starts: %s", self.server.errorString())
                return
        self.server.newConnection.connect(lambda: self._accept(on_paths))

    def _accept(self, on_paths: Callable[[list[str]], None]) -> None:
        while self.server.hasPendingConnections():
            self._serve(self.server.nextPendingConnection(), on_paths)

    def _serve(self, socket: QLocalSocket, on_paths: Callable[[list[str]], None]) -> None:
        buffer = bytearray()
        handled = []

        def read():
            if handled:
                return
            buffer.extend(bytes(socket.readAll()))
            if not buffer.endswith(b"\n"):
                return
            handled.append(True)
            try:
                paths = json.loads(buffer.decode("utf-8"))
            except ValueError:
                paths = []
            on_paths([p for p in paths if isinstance(p, str)])
            socket.disconnectFromServer()

        socket.readyRead.connect(read)
        socket.disconnected.connect(read)            # the data can arrive together with the goodbye
        socket.disconnected.connect(socket.deleteLater)
        read()                                       # or before we started listening for it
