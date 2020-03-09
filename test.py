import sys
import mmap
import json

def main():
    if len(sys.argv) < 5:
        return
    param= v2py(sys.argv[1], int(sys.argv[2]))
    py2v(func(param), sys.argv[3], int(sys.argv[4]))


def v2py(name_, size_):
    with mmap.mmap(-1, size_, name_) as m:
        val = m.read().decode()    
        return json.loads(val)


def py2v(data_, name_, size_):
    expr = json.dumps(data_).encode()
    if len(expr) <= size_:
        with mmap.mmap(-1, len(expr), name_) as m:
            m.write(expr)
            return True
    else:
        return False


def func(m):
    return [2.13434 * x for x in m]


if __name__ == '__main__':
    main()
