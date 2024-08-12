import time

class Timer:
    def __init__(self):
        self.start()
    def start(self):
        self.start = time.perf_counter()
    def check(self):
        t = time.perf_counter() - self.start
        return "(🕑 {:.2f}ms)".format(t * 1000) if t < 1 else "(🕑 {:.2f}s)".format(t)