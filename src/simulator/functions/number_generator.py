from numpy import random as np_random

def generate_random_numbers(n, seed=None,min=1,max=1.3):
    """
    Generate n random numbers using the seed provided.
    """
    if seed is not None:
        np_random.seed(seed)
    return np_random.uniform(min,max,n)