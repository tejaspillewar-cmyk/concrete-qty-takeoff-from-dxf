"""
Geometry utilities for polygon calculations.
Pure math — no DXF dependency.
"""
import math

def polygon_area_sqmm(vertices: list[tuple[float, float]]) -> float:
    """
    Compute area of a polygon using the Shoelace formula.
    
    Args:
        vertices: List of (x, y) tuples in mm.
    
    Returns:
        Absolute area in square millimeters.
    """
    n = len(vertices)
    if n < 3:
        return 0.0
    area = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def sqmm_to_sqm(area_sqmm: float) -> float:
    """Convert square millimeters to square meters."""
    return area_sqmm / 1_000_000.0


def polygon_centroid(vertices: list[tuple[float, float]]) -> tuple[float, float]:
    """
    Compute the geometric centroid of a polygon.
    
    Args:
        vertices: List of (x, y) tuples.
    
    Returns:
        (cx, cy) centroid coordinates.
    """
    n = len(vertices)
    if n == 0:
        return (0.0, 0.0)
    cx = sum(v[0] for v in vertices) / n
    cy = sum(v[1] for v in vertices) / n
    return (cx, cy)


def point_in_polygon(px: float, py: float, polygon: list[tuple[float, float]]) -> bool:
    """
    Test if a point (px, py) is inside a polygon using the ray-casting algorithm.
    
    Args:
        px, py: Point coordinates.
        polygon: List of (x, y) vertex tuples.
    
    Returns:
        True if point is inside the polygon.
    """
    n = len(polygon)
    if n < 3:
        return False
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def polygon_bbox(vertices: list[tuple[float, float]]) -> tuple[float, float, float, float]:
    """
    Compute the bounding box of a polygon.
    
    Returns:
        (min_x, min_y, max_x, max_y)
    """
    xs = [v[0] for v in vertices]
    ys = [v[1] for v in vertices]
    return (min(xs), min(ys), max(xs), max(ys))


# ──────────────────────────────────────────────────────────────────────────────
# Beam Math (Lines & Intersections)
# ──────────────────────────────────────────────────────────────────────────────

def distance(p1: tuple[float, float], p2: tuple[float, float]) -> float:
    """Distance between two points."""
    return math.hypot(p1[0] - p2[0], p1[1] - p2[1])


def line_angle(p1: tuple[float, float], p2: tuple[float, float]) -> float:
    """Returns angle of line in degrees, normalized to [0, 180)."""
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    angle = math.degrees(math.atan2(dy, dx))
    if angle < 0:
        angle += 180.0
    elif angle >= 180.0:
        angle -= 180.0
    return angle


def point_line_distance(pt: tuple[float, float], lp1: tuple[float, float], lp2: tuple[float, float]) -> float:
    """Perpendicular distance from point pt to infinite line defined by lp1, lp2."""
    num = abs((lp2[1] - lp1[1])*pt[0] - (lp2[0] - lp1[0])*pt[1] + lp2[0]*lp1[1] - lp2[1]*lp1[0])
    den = distance(lp1, lp2)
    if den == 0:
        return distance(pt, lp1)
    return num / den


def are_lines_parallel(l1p1, l1p2, l2p1, l2p2, tol_deg=2.0) -> bool:
    """Check if two lines are parallel within a tolerance."""
    a1 = line_angle(l1p1, l1p2)
    a2 = line_angle(l2p1, l2p2)
    diff = abs(a1 - a2)
    return diff < tol_deg or abs(diff - 180.0) < tol_deg


def project_point_on_line(pt: tuple[float, float], lp1: tuple[float, float], lp2: tuple[float, float]) -> float:
    """
    Project point pt onto line (lp1, lp2).
    Returns the parameter t where projected point = lp1 + t*(lp2 - lp1).
    """
    dx = lp2[0] - lp1[0]
    dy = lp2[1] - lp1[1]
    L2 = dx*dx + dy*dy
    if L2 == 0:
        return 0.0
    return ((pt[0] - lp1[0])*dx + (pt[1] - lp1[1])*dy) / L2


def check_lines_overlap(l1p1, l1p2, l2p1, l2p2) -> float:
    """
    Check if parallel line 2 overlaps with line 1 longitudinally.
    Returns the overlapping length in mm. If no overlap, returns 0.
    """
    # Project l2 endpoints onto l1
    t1 = project_point_on_line(l2p1, l1p1, l1p2)
    t2 = project_point_on_line(l2p2, l1p1, l1p2)
    
    t_min = min(t1, t2)
    t_max = max(t1, t2)
    
    # Overlap interval on l1 is [max(0, t_min), min(1, t_max)]
    overlap_start = max(0.0, t_min)
    overlap_end = min(1.0, t_max)
    
    if overlap_start < overlap_end:
        # Calculate physical overlap length
        L = distance(l1p1, l1p2)
        return (overlap_end - overlap_start) * L
    return 0.0


def distance_pt_to_polygon(pt: tuple[float, float], poly: list[tuple[float, float]]) -> float:
    """Find the shortest distance from a point to any edge of a polygon."""
    min_dist = float('inf')
    n = len(poly)
    for i in range(n):
        p1 = poly[i]
        p2 = poly[(i+1)%n]
        # Point to segment distance
        L2 = (p2[0]-p1[0])**2 + (p2[1]-p1[1])**2
        if L2 == 0:
            d = distance(pt, p1)
        else:
            t = max(0, min(1, ((pt[0]-p1[0])*(p2[0]-p1[0]) + (pt[1]-p1[1])*(p2[1]-p1[1])) / L2))
            proj = (p1[0] + t*(p2[0]-p1[0]), p1[1] + t*(p2[1]-p1[1]))
            d = distance(pt, proj)
        if d < min_dist:
            min_dist = d
    return min_dist
