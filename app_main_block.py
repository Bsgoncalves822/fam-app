if __name__ == '__main__':
    # Auto-update from GitHub before starting
    try:
        from patch_update import run_update
        run_update(verbose=False)
    except Exception as e:
        logger.warning(f"Auto-update falhou (continuando): {e}")

    init_db()
    migrate_csvs_to_db()
    cfg = load_config()
    logger.info(f"FAM App starting on port {cfg['port']}")
    app.run(host='0.0.0.0', port=cfg['port'], debug=False)
