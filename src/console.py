def info(message: str, logger) -> None:
    """
    Ausgabe auf Konsole und Logdatei.
    """
    print(message)
    logger.info(message)