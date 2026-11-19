""" Version of the OXML Filler, answered by `--version` and by the `version` request of the embedded runtime """

__version__ = "1.1.0"

# version of the contract of requests and answers shared by the command line and the embedded runtime;
# bump it when a request or an answer changes in a way the app has to know
PROTOCOL = 1


def describe_version() -> dict:
    return {"ok": True, "version": __version__, "protocol": PROTOCOL}
