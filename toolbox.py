from fuzzylogic.classes import Domain
from fuzzylogic.functions import R, S, trapezoid

def pointArrangement(minVal, maxVal):
    points = [0] * 8
    points[0] = minVal
    points[1] = maxVal / 5
    points[2] = points[1] + minVal
    points[3] = points[1] * 2
    points[4] = points[3] + minVal
    points[5] = points[1] * 3
    points[6] = points[5] + minVal
    points[7] = points[1] * 4
    return points

def generate_domain(name, minVal, maxVal, labels, res=0.1):
    points = pointArrangement(minVal, maxVal)
    domain = Domain(name, 0, maxVal, res=res)
    # Label-label pertama selalu S, tengah trapezoid, terakhir R
    setattr(domain, labels[0], S(points[0], points[1]))
    setattr(domain, labels[1], trapezoid(points[0], points[1], points[2], points[3], c_m=1))
    setattr(domain, labels[2], trapezoid(points[2], points[3], points[4], points[5], c_m=1))
    setattr(domain, labels[3], trapezoid(points[4], points[5], points[6], points[7], c_m=1))
    setattr(domain, labels[4], R(points[6], points[7]))
    return domain