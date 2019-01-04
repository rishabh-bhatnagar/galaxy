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
    :param operations:
    """
    for i in operations:
        print(i())


def add(n1, n2):
    return n1 + n2


def sub(n1, n2):
    return n1 - n2


print(which_is_better(lambda: add(1,2),lambda: sub(1,2)))
