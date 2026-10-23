"""Full launcher window: both 'Scan Archive' and 'Archive Email' buttons."""
from email_archiver.config import load_config, setup_logging
from email_archiver.ui.app import LauncherApp


def main() -> None:
    cfg = load_config()
    setup_logging(cfg)
    LauncherApp(cfg).run()
