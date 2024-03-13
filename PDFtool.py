from panel import Menu


def run():
    try:
        menu = Menu()
        menu.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    run()
