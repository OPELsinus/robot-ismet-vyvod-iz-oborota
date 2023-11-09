def get_hostname() -> str:
    import socket
    return socket.gethostbyname(socket.gethostname())
