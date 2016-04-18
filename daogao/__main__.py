from daogao.server import main_loop

try:
    main_loop()
except KeyboardInterrupt:
    print("bye")