from time import time_ns as time
import sys


sys.maxsize = 10**70
def get_run_time(function, *args, **kwargs):
    t1 = time()
    function(*args, **kwargs)
    t2 = time()
    return t2-t1


def which_is_better(*operations: list) -> int:
    """
    This function returns index of operation which takes the least time to execute
    there are two versions of this method.
    1) Which takes function and parameters as a two tuple
        for example:
            def add(n1, n2):
                return n1+n2
            def sub(n1, n2):
                return n1-n2

            this will be passed as:
            which_is_better( (add, (1,2)), (sub, (1,2) ) )
    2) Which takes executable lambda which takes no parameters and executes the required function on called.
        for example:
            with functions defined same as of above cases' functions:
                which_is_better(
                    lambda: add(1,2),
                    lambda: sub(1,2)
                )

    :param operations: either a tuple having two params that is function and params or
                              a function whicn will be called without any parameters.
    """

    running_times = [get_run_time(operation[0], operation[1]) if isinstance(operation, tuple) else get_run_time(operation) for operation in operations]
    print([i/10**0 for i in running_times])


ele = [i for i in range(10**6)]


def function1():
    [int(i) for i in ele]


def function2():
    var = (int(i) for i in ele)


which_is_better(function1, function2)
