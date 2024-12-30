def parse_txt_to_data(filename):
    data = []  # To store the parsed data

    with open(filename, 'r') as file:
        content = file.read().strip()  # Read the entire content and remove leading/trailing whitespace
        
        # Split the content based on the double newline that separates different groups of data
        groups = content.split('\n\n')
        
        # Loop through each group of data
        for group in groups:
            records = []
            lines = group.split('\n')  # Split by single newline to get each line

            # For each line in the group, convert to a dictionary
            for line in lines:
                # Split each line by the comma
                parts = line.split(',')
                
                # Create a dictionary for each line
                record = {
                    "khedma": parts[0].strip(),
                    "qte": parts[1].strip(),
                    "Prix": parts[2].strip()
                }
                
                # Add the record to the records list
                records.append(record)

            # Add the records list (group of data) to the main data list
            data.append(records)

    return data