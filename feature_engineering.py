import numpy as np
import pandas as pd

def calculate_angle(a, b, c):
    """Calculate 3D angle between three joints (a -> b -> c) where b is the vertex."""
    a = np.array(a[:3])
    b = np.array(b[:3])
    c = np.array(c[:3])
    
    ba = a - b
    bc = c - b
    
    cosine_angle = np.dot(ba, bc) / (np.linalg.norm(ba) * np.linalg.norm(bc) + 1e-6)
    angle = np.arccos(np.clip(cosine_angle, -1.0, 1.0))
    return float(np.degrees(angle))

def calculate_angle_2d(a, b, c):
    """Calculate 2D planar angle between three joints (a -> b -> c) where b is the vertex."""
    ba = np.array([a[0] - b[0], a[1] - b[1]])
    bc = np.array([c[0] - b[0], c[1] - b[1]])
    norm_ba = np.linalg.norm(ba)
    norm_bc = np.linalg.norm(bc)
    if norm_ba < 1e-6 or norm_bc < 1e-6:
        return 180.0
    cosine = np.dot(ba, bc) / (norm_ba * norm_bc)
    angle = np.arccos(np.clip(cosine, -1.0, 1.0))
    return float(np.degrees(angle))

def calculate_torso_inclination(shoulder, hip):
    """
    Calculate torso inclination angle from vertical (0 deg = perfectly upright).
    In image space, hip is at (hx, hy), shoulder is at (sx, sy).
    """
    dx = abs(shoulder[0] - hip[0])
    dy = abs(shoulder[1] - hip[1])
    if dy < 1e-6:
        return 90.0
    return float(np.degrees(np.arctan2(dx, dy)))

def calculate_hip_deviation_2d(shoulder, hip, ankle):
    """
    Computes signed vertical deviation of the hip from the shoulder-ankle baseline.
    Positive value indicates hips are sagging downward towards floor.
    Negative value indicates hips are piked upward towards ceiling.
    """
    sa = np.array([ankle[0] - shoulder[0], ankle[1] - shoulder[1]])
    norm_sa = np.linalg.norm(sa)
    if norm_sa < 1e-6:
        return 0.0
    u = sa / norm_sa
    sh = np.array([hip[0] - shoulder[0], hip[1] - shoulder[1]])
    proj_dist = np.dot(sh, u)
    proj_pt = np.array([shoulder[0], shoulder[1]]) + proj_dist * u
    # In screen coordinates, y increases downwards.
    return float(hip[1] - proj_pt[1])

def extract_features_from_landmarks(landmarks_list):
    """
    Given a list of 33 MediaPipe landmarks [[x,y,z], ...], 
    returns a dictionary of engineered biomechanical features.
    """
    # MediaPipe Keypoint Index Map:
    # Left Shoulder: 11, Right Shoulder: 12
    # Left Elbow: 13, Right Elbow: 14
    # Left Wrist: 15, Right Wrist: 16
    # Left Hip: 23, Right Hip: 24
    # Left Knee: 25, Right Knee: 26
    # Left Ankle: 27, Right Ankle: 28
    
    # 1. Hip Center Normalization (Translation Invariance)
    hip_center = (np.array(landmarks_list[23][:3]) + np.array(landmarks_list[24][:3])) / 2.0
    normalized_landmarks = []
    for lm in landmarks_list:
        norm_lm = np.array(lm[:3]) - hip_center
        normalized_landmarks.extend(norm_lm.tolist())
    
    # 2. Biomechanical Joint Angles
    knee_angle_left = calculate_angle(landmarks_list[23], landmarks_list[25], landmarks_list[27])
    knee_angle_right = calculate_angle(landmarks_list[24], landmarks_list[26], landmarks_list[28])
    
    elbow_angle_left = calculate_angle(landmarks_list[11], landmarks_list[13], landmarks_list[15])
    elbow_angle_right = calculate_angle(landmarks_list[12], landmarks_list[14], landmarks_list[16])
    
    back_angle_left = calculate_angle(landmarks_list[11], landmarks_list[23], landmarks_list[25])
    back_angle_right = calculate_angle(landmarks_list[12], landmarks_list[24], landmarks_list[26])
    
    hip_angle_left = calculate_angle(landmarks_list[11], landmarks_list[23], landmarks_list[27])
    hip_angle_right = calculate_angle(landmarks_list[12], landmarks_list[24], landmarks_list[28])
    
    shoulder_angle_left = calculate_angle(landmarks_list[23], landmarks_list[11], landmarks_list[13])
    shoulder_angle_right = calculate_angle(landmarks_list[24], landmarks_list[12], landmarks_list[14])
    
    # 3. Key Joint Distance Ratios (Scale Invariant)
    torso_height = np.linalg.norm(np.array(landmarks_list[11][:3]) - np.array(landmarks_list[23][:3])) + 1e-6
    wrist_dist = np.linalg.norm(np.array(landmarks_list[15][:3]) - np.array(landmarks_list[16][:3])) / torso_height
    ankle_dist = np.linalg.norm(np.array(landmarks_list[27][:3]) - np.array(landmarks_list[28][:3])) / torso_height
    knee_dist = np.linalg.norm(np.array(landmarks_list[25][:3]) - np.array(landmarks_list[26][:3])) / torso_height
    
    features = {
        'knee_angle_left': knee_angle_left,
        'knee_angle_right': knee_angle_right,
        'elbow_angle_left': elbow_angle_left,
        'elbow_angle_right': elbow_angle_right,
        'back_angle_left': back_angle_left,
        'back_angle_right': back_angle_right,
        'hip_angle_left': hip_angle_left,
        'hip_angle_right': hip_angle_right,
        'shoulder_angle_left': shoulder_angle_left,
        'shoulder_angle_right': shoulder_angle_right,
        'wrist_dist_ratio': wrist_dist,
        'ankle_dist_ratio': ankle_dist,
        'knee_dist_ratio': knee_dist,
    }
    
    # Combine normalized coordinates with calculated features
    for idx, val in enumerate(normalized_landmarks):
        features[f'norm_coord_{idx}'] = val
        
    return features

def process_csv_dataframe(df):
    """
    Transforms raw coordinate DataFrame into an enriched feature DataFrame.
    """
    feature_rows = []
    
    for idx, row in df.iterrows():
        landmarks_list = []
        for i in range(33):
            landmarks_list.append([row[f'x{i}'], row[f'y{i}'], row[f'z{i}']])
            
        feats = extract_features_from_landmarks(landmarks_list)
        feature_rows.append(feats)
        
    feature_df = pd.DataFrame(feature_rows)
    if 'exercise_name' in df.columns:
        feature_df['exercise_name'] = df['exercise_name'].values
        
    return feature_df
