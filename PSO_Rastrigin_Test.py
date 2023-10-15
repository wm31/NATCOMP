import numpy as np
from NAT_cw_a import PSO



def test_PSO():
    print("Starting the PSO test...")  # Indicate the start of the test
    # Define test parameters
    num_particles = 10
    dimensions = 2
    num_iterations = 100
    target_error = 1e-5
    c1 = 2.0
    c2 = 2.0
    w = 0.7
    bounds = [(-5.12, 5.12)] * dimensions

    # Execute PSO algorithm
    gbest_position, gbest_score = PSO(num_particles, dimensions, num_iterations, target_error, c1, c2, w, bounds)

    # Print the results
    print(f"Best position found by PSO: {gbest_position}")
    print(f"Best score found by PSO: {gbest_score}")

    # Assert if the solution is within bounds
    for i in range(dimensions):
        assert bounds[i][0] <= gbest_position[i] <= bounds[i][1], f"Position {i} out of bounds."

    # Assert if the solution is close to the global minimum
    assert np.abs(gbest_score - 0.0) < target_error, f"Score {gbest_score} not close enough to global minimum."

    print("PSO test completed successfully!")  # Indicate the successful completion of the test

test_PSO()