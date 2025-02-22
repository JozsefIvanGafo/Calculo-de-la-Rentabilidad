from simulator import MonteCarloSimulator
from parameters import Parameters


if __name__ == "__main__":
    simulator = MonteCarloSimulator(Parameters())
    simulator.run()
    simulator.plot()