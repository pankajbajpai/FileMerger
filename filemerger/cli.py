import socket
import sys
import threading
import webbrowser

from .app import create_app


def _find_free_port(preferred=5678):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", preferred))
        s.close()
        return preferred
    except OSError:
        s.close()
        s2 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s2.bind(("127.0.0.1", 0))
        port = s2.getsockname()[1]
        s2.close()
        return port


def main():
    app = create_app()
    port = _find_free_port()
    url = f"http://127.0.0.1:{port}"

    no_browser = "--no-browser" in sys.argv

    if not no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    print(f"File Merger is running at {url}")
    print("Open that address in your browser if it did not open automatically.")
    print("Press CTRL+C to stop.")

    app.run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
