NUM_ROWS =25
NUM_COLS = 25

my_matrix = []
for row in range(NUM_ROWS):
    new_row = []
    for col in range(NUM_COLS):
        new_row.append(row * col)
    my_matrix.append(new_row)
    
trace = 0
for rc in range(0,NUM_COLS):
     trace += my_matrix[rc][rc]
print(trace)
