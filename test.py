import sys
import mmap
import json

def main():
    if len(sys.argv) < 3:
        return
    param= v2py(sys.argv[1])
    py2v(func(param), sys.argv[2])


def v2py(name_size):
    name, size = name_size.split('+')
    with mmap.mmap(-1, int(size), name) as m:
        val = m.read().decode()
        return json.loads(val)


def py2v(data_, name_size):
    expr = json.dumps(data_).encode()
    name, size = name_size.split('+')
    if len(expr) <= int(size):
        with mmap.mmap(-1, len(expr), name) as m:
            m.write(expr)
            return True
    else:
        return False


def func(m):
    return [2.13434 * x for x in m]


if __name__ == '__main__':
    main()
