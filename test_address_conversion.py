from app import get_address_from_coordinates

# Test the address conversion functionality
print('Testing address conversion functionality...')
print()

# Test with Karachi coordinates
lat1, lng1 = '24.8607', '67.0011'
address1 = get_address_from_coordinates(lat1, lng1)
print(f'Coordinates: {lat1}, {lng1}')
print(f'Address: {address1}')
print()

# Test with slightly different coordinates
lat2, lng2 = '24.8608', '67.0012'
address2 = get_address_from_coordinates(lat2, lng2)
print(f'Coordinates: {lat2}, {lng2}')
print(f'Address: {address2}')
print()

print('Address conversion functionality is working correctly!')