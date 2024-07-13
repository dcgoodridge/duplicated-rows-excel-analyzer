import pandas as pd
import random
from faker import Faker

fake = Faker()


def generate_similar_data(data_type, similarity_probability):
    data_map = {
        'money': lambda: f"${random.randint(1, 1000):,.2f}",
        'street_address': fake.street_address,
        'postal_code': fake.postcode,
        'company_name': fake.company,
        'random_word': fake.word
    }

    data = []
    unique_data = set()

    for _ in range(1000):  # Generate a pool of unique data points
        value = data_map[data_type]()
        if value not in unique_data:
            unique_data.add(value)

    unique_data = list(unique_data)
    num_similar_items = int(len(unique_data) * similarity_probability / 100)

    data.extend(random.choices(unique_data, k=num_similar_items))

    for _ in range(len(unique_data) - num_similar_items):
        value = random.choice(unique_data)
        value = modify_value(value, data_type)
        data.append(value)

    random.shuffle(data)
    return data


def modify_value(value, data_type):
    if data_type == 'money':
        value = f"${float(value[1:].replace(',', '')) + random.uniform(-5, 5):,.2f}"
    else:
        # Randomly append a word
        value += fake.word()

        # Randomly switch characters
        value = switch_characters(value)

        # Randomly change vowels to vowels with accents
        value = change_vowels_with_accents(value)

    return value


def switch_characters(value):
    if len(value) > 1:
        i = random.randint(0, len(value) - 2)
        value = value[:i] + value[i + 1] + value[i] + value[i + 2:]
    return value


def change_vowels_with_accents(value):
    accents = {
        'a': ['á', 'à', 'â', 'ä'],
        'e': ['é', 'è', 'ê', 'ë'],
        'i': ['í', 'ì', 'î', 'ï'],
        'o': ['ó', 'ò', 'ô', 'ö'],
        'u': ['ú', 'ù', 'û', 'ü']
    }

    for vowel, accented_vowels in accents.items():
        if vowel in value:
            value = value.replace(vowel, random.choice(accented_vowels))

    return value


def generate_excel(num_rows, num_columns, data_types, similarity_probability, output_path):
    data = {}

    for col in range(num_columns):
        data_type = data_types[col % len(data_types)]
        column_data = generate_similar_data(data_type, similarity_probability)
        data[f'Column {col + 1}'] = random.choices(column_data, k=num_rows)

    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)


# Example usage
num_rows = 1000
num_columns = 5
data_types = ['street_address', 'postal_code', 'company_name', 'random_word', 'money']
similarity_probability = 50  # Probability of similarity in percentage
output_path = 'test_data.xlsx'

generate_excel(num_rows, num_columns, data_types, similarity_probability, output_path)
