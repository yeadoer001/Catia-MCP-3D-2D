import sys
from .server import CATIAMCPServer

def main():
    server = CATIAMCPServer()
    try:
        server.setup()
        server.run()
    except Exception as e:
        print(f"Failed to start CATIA MCP Server: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()