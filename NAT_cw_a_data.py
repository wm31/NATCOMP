    import xlwt
    from datetime import datetime
    import numpy as np
    from PSO_algorithm import PSO

    # Define the range of population sizes and dimensions to test
    population_sizes = [10, 20, 30, 40, 50]
    dimensions = [2, 5, 10, 20, 50]

    # Define the PSO parameters
    num_iterations = 1000
    target_error = 1e-6
    c1 = 2.0
    c2 = 2.0
    w = 0.7
    bounds = [(-5.12, 5.12)] * max(dimensions)

    # Create a new workbook and worksheet
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("PSO Results")

    # Write the headings to the worksheet
    worksheet.write(0, 0, "Population Size")
    worksheet.write(0, 1, "Dimensions")
    worksheet.write(0, 2, "Best Score")
    worksheet.write(0, 3, "Best Position")

    # Run the PSO algorithm for each combination of population size and dimensions
    row = 1
    for num_particles in population_sizes:
        for d in dimensions:
            start_time = datetime.now()
            best_position, best_score = PSO(num_particles, d, num_iterations, target_error, c1, c2, w, bounds)
            end_time = datetime.now()

            # Write the results to the worksheet
            worksheet.write(row, 0, num_particles)
            worksheet.write(row, 1, d)
            worksheet.write(row, 2, best_score)
            worksheet.write(row, 3, str(best_position))
            worksheet.write(row, 4, str(end_time - start_time))

            row += 1

    # Save the workbook
    workbook.save("PSO_Results.xls")