import cv2
import numpy as np


def auto_crop_smart(img):
    """
    Smart Cropping: Removes static black borders using morphological operations.
    """
    if img is None: return None
    try:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 15, 255, cv2.THRESH_BINARY)

        # Morphological opening to remove noise
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

        coords = cv2.findNonZero(thresh)
        if coords is not None:
            x, y, w, h = cv2.boundingRect(coords)
            return img[y:y + h, x:x + w]
        return img
    except Exception:
        return img


def get_blur_score(img):
    """
    Blur Detection: Normalized Laplacian Variance.
    Result is independent of image resolution.
    """
    if img is None: return 0
    try:
        h, w = img.shape[:2]
        scale = 500 / w
        dim = (500, int(h * scale))
        resized = cv2.resize(img, dim, interpolation=cv2.INTER_AREA)

        gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
        return cv2.Laplacian(gray, cv2.CV_64F).var()
    except Exception:
        return 0


def get_frame_diff(img1, img2):
    """
    MSE Calculation: Fast frame difference metric.
    Returns: Float (0.0 - 100.0+), lower is more similar.
    """
    if img1 is None or img2 is None: return 100.0
    try:
        # Convert to grayscale and resize for speed
        g1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        g2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)

        g1 = cv2.resize(g1, (64, 64))
        g2 = cv2.resize(g2, (64, 64))

        # Calculate Mean Squared Error
        err = np.sum((g1.astype("float") - g2.astype("float")) ** 2)
        err /= float(g1.shape[0] * g1.shape[1])

        return err / 100.0  # Normalized roughly
    except Exception:
        return 100.0


def get_dhash(img):
    """
    Difference Hash: Structural fingerprinting.
    Returns: 64-bit boolean array.
    """
    try:
        resized = cv2.resize(img, (9, 8), interpolation=cv2.INTER_AREA)
        gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
        return gray[:, 1:] > gray[:, :-1]
    except Exception:
        return None


def hamming_distance(hash1, hash2):
    """Compare two dHash fingerprints."""
    if hash1 is None or hash2 is None: return 64
    return np.count_nonzero(hash1 != hash2)