import cv2
import mediapipe as mp
import os
import csv

# Initialize MediaPipe Pose (Use static_image_mode=True for images)
mp_pose = mp.solutions.pose
pose = mp_pose.Pose(static_image_mode=True, min_detection_confidence=0.5)

# List of exercise folder names
folders = ['push up', 'pull up', 'squat', 'deadlift', 'plank']
all_rows = []

print("Processing started... Please wait...")

for folder_name in folders:
    # Create the folder path (Example: dataset/pushup)
    folder_path = os.path.join("IMAGE DATASET", folder_name)

    # Check if the folder exists
    if not os.path.exists(folder_path):
        print(f"Warning: '{folder_name}' folder not found!")
        continue

    print(f"Reading images from '{folder_name}' folder...")

    # Loop through each image in the folder
    for image_name in os.listdir(folder_path):
        image_path = f"{folder_path}/{image_name}"

        # Read the image
        image = cv2.imread(image_path)
        if image is None:
            continue  # Skip if the image is corrupted

        # Convert the image from BGR to RGB for MediaPipe
        image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        results = pose.process(image_rgb)

        # Process only if a person is detected
        if results.pose_landmarks:
            single_image_data = []

            # Extract X, Y, Z coordinates of all 33 pose landmarks
            for lm in results.pose_landmarks.landmark:
                single_image_data.append(lm.x)
                single_image_data.append(lm.y)
                single_image_data.append(lm.z)

            # Add the exercise name as the label
            single_image_data.append(folder_name)

            # Store the data
            all_rows.append(single_image_data)

# Create the CSV file
header = []
for i in range(33):
    header.extend([f'x{i}', f'y{i}', f'z{i}'])
header.append('exercise_name')  # Label column

with open('fitness_data.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(header)
    writer.writerows(all_rows)

print(f"\nSuccess! Data from {len(all_rows)} images has been saved to 'fitness_data.csv'.")