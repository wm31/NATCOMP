import numpy as np
import openpyxl

# Define the Rastrigin function as an example objective function
def rastrigin(x):
    d = len(x)
    return 10 * d + np.sum(x ** 2 - 10 * np.cos(2 * np.pi * x))

# Particle class definition
class Particle:
    def __init__(self, dimensions, bounds):
        self.position = np.random.uniform(bounds[0], bounds[1], dimensions)
        self.velocity = np.random.uniform(-0.5, 0.5, dimensions)
        self.best_position = np.copy(self.position)
        self.best_score = np.inf

# PSO Algorithm
def PSO(num_particles, dimensions, num_iterations, target_error, cognitive_coeff, social_coeff, inertia_weight, bounds, objective_function):
    particles = [Particle(dimensions, bounds) for _ in range(num_particles)]
    gbest_position = np.random.uniform(bounds[0], bounds[1], dimensions)
    gbest_score = np.inf

    # Initialization
    for particle in particles:
        current_score = objective_function(particle.position)
        if current_score < particle.best_score:
            particle.best_position = np.copy(particle.position)
            particle.best_score = current_score
        if current_score < gbest_score:
            gbest_position = np.copy(particle.position)
            gbest_score = current_score

    # Main loop
    for iteration in range(num_iterations):
        for particle in particles:
            # Update velocity and position
            inertia = inertia_weight * particle.velocity
            personal_attraction = cognitive_coeff * np.random.random() * (particle.best_position - particle.position)
            social_attraction = social_coeff * np.random.random() * (gbest_position - particle.position)
            particle.velocity = inertia + personal_attraction + social_attraction
            particle.position = np.clip(particle.position + particle.velocity, bounds[0], bounds[1])
            
            # Update best scores
            current_score = objective_function(particle.position)
            if current_score < particle.best_score:
                particle.best_position = np.copy(particle.position)
                particle.best_score = current_score
            if current_score < gbest_score:
                gbest_position = np.copy(particle.position)
                gbest_score = current_score

        # Stopping criterion
        if gbest_score < target_error:
            break

    return gbest_position, gbest_score

def save_to_excel(results, filename="PSO_Results.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PSO Results"
    ws.append(["Population Size", "Dimension Size", "Best Score", "Best Position"])
    for result in results:
        ws.append(result)
    wb.save(filename)

# Test the PSO algorithm
pop_sizes = [10, 20, 30]
dim_sizes = [2, 5, 10]
num_iterations = 100
target_error = 1e-6
cognitive_coeff = 2.0
social_coeff = 2.0
inertia_weight = 0.7
bounds = (-5.12, 5.12)

results = []
for pop_size in pop_sizes:
    for dim_size in dim_sizes:
        best_pos, best_score = PSO(pop_size, dim_size, num_iterations, target_error, cognitive_coeff, social_coeff, inertia_weight, bounds, rastrigin)
        results.append([pop_size, dim_size, best_score, str(best_pos)])

save_to_excel(results)
