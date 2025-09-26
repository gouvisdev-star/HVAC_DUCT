using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVAC_DUCT
{
    public class Supply_duct
    {
        /// <summary>
        /// DrawableOverrule để vẽ thêm các tick vuông góc với Polyline (không có fillet arc)
        /// </summary>
        public class FilletOverrule : DrawableOverrule
        {
            private double _blueWidth = 8.0; // Độ dài đoạn xanh
            private static System.Collections.Generic.HashSet<ObjectId> _allowedPolylines = new System.Collections.Generic.HashSet<ObjectId>();

            /// <summary>
            /// Constructor khởi tạo bộ lọc tùy chỉnh
            /// </summary>
            public FilletOverrule()
            {
                SetCustomFilter(); // Gọi bộ lọc riêng
            }

            /// <summary>
            /// Cập nhật độ dài đoạn xanh
            /// </summary>
            /// <param name="newWidth">Độ dài mới cho đoạn xanh</param>
            public void UpdateBlueWidth(double newWidth)
            {
                _blueWidth = newWidth;
            }

            /// <summary>
            /// Thêm polyline vào danh sách được phép hiển thị tick
            /// </summary>
            /// <param name="polylineId">ObjectId của polyline</param>
            public static void AddAllowedPolyline(ObjectId polylineId)
            {
                _allowedPolylines.Add(polylineId);
            }

            /// <summary>
            /// Xóa tất cả polyline khỏi danh sách
            /// </summary>
            public static void ClearAllowedPolylines()
            {
                _allowedPolylines.Clear();
            }

            /// <summary>
            /// Vẽ polyline fit tạm màu vàng
            /// </summary>
            /// <param name="pl">Polyline gốc</param>
            /// <param name="wd">WorldDraw context</param>
            private void DrawTempPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, WorldDraw wd)
            {
                try
                {
                    if (pl.NumberOfVertices < 2) return;

                    // Vẽ polyline fit tạm màu vàng
                    wd.SubEntityTraits.Color = 2; // Vàng
                    
                    // Tạo danh sách điểm cho polyline fit
                    List<Point3d> fitPoints = new List<Point3d>();
                    
                    // Thêm tất cả các đỉnh của polyline
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(i));
                    }
                    
                    // Nếu polyline đóng, thêm điểm đầu để tạo polyline fit đóng
                    if (pl.Closed && pl.NumberOfVertices > 2)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(0));
                    }
                    
                    // Vẽ polyline fit qua các điểm
                    if (fitPoints.Count >= 2)
                    {
                        // Vẽ polyline fit bằng cách vẽ nhiều line ngắn để tạo đường cong mượt
                        DrawPolylineFit(fitPoints, wd);
                    }
                }
                catch
                {
                    // Bỏ qua lỗi nếu có
                }
            }

            /// <summary>
            /// Vẽ Polyline Fit mượt mà qua các điểm
            /// </summary>
            /// <param name="points">Danh sách điểm</param>
            /// <param name="wd">WorldDraw context</param>
            private void DrawPolylineFit(List<Point3d> points, WorldDraw wd)
            {
                try
                {
                    if (points.Count < 2) return;

                    // Tạo Polyline Fit với các điểm trung gian để tạo đường cong mượt
                    List<Point3d> fitPoints = new List<Point3d>();
                    
                    // Thêm điểm đầu
                    fitPoints.Add(points[0]);
                    
                    // Tạo các điểm trung gian cho Polyline Fit
                    for (int i = 0; i < points.Count - 1; i++)
                    {
                        Point3d p1 = points[i];
                        Point3d p2 = points[i + 1];
                        
                        // Tạo các điểm trung gian để tạo đường cong mượt
                        int segments = 6; // Số đoạn giữa 2 điểm
                        for (int j = 1; j < segments; j++)
                        {
                            double t = (double)j / segments;
                            
                            // Sử dụng interpolation đơn giản để tạo đường cong mượt
                            Point3d smoothPoint = new Point3d(
                                p1.X + (p2.X - p1.X) * t,
                                p1.Y + (p2.Y - p1.Y) * t,
                                p1.Z + (p2.Z - p1.Z) * t
                            );
                            fitPoints.Add(smoothPoint);
                        }
                    }
                    
                    // Thêm điểm cuối
                    fitPoints.Add(points[points.Count - 1]);
                    
                    // Vẽ Polyline Fit bằng các line ngắn
                    for (int i = 0; i < fitPoints.Count - 1; i++)
                    {
                        wd.Geometry.WorldLine(fitPoints[i], fitPoints[i + 1]);
                    }
                }
                catch
                {
                    // Bỏ qua lỗi nếu có
                }
            }

            /// <summary>
            /// Lấy điểm trên polyline tạm tại khoảng cách cho trước
            /// </summary>
            /// <param name="pl">Polyline gốc</param>
            /// <param name="distance">Khoảng cách</param>
            /// <returns>Điểm trên polyline tạm</returns>
            private Point3d GetPointOnTempPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double distance)
            {
                try
                {
                    // Tạo danh sách điểm cho polyline fit
                    List<Point3d> fitPoints = new List<Point3d>();
                    
                    // Thêm tất cả các đỉnh của polyline
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(i));
                    }
                    
                    // Nếu polyline đóng, thêm điểm đầu
                    if (pl.Closed && pl.NumberOfVertices > 2)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(0));
                    }
                    
                    // Tính tổng chiều dài polyline fit
                    double totalLength = 0.0;
                    for (int i = 0; i < fitPoints.Count - 1; i++)
                    {
                        totalLength += fitPoints[i].DistanceTo(fitPoints[i + 1]);
                    }
                    
                    if (totalLength <= 0) return pl.GetPointAtDist(distance);
                    
                    // Chuẩn hóa distance
                        if (distance > totalLength) distance = totalLength;

                    // Tìm điểm trên polyline fit
                    double currentLength = 0.0;
                    for (int i = 0; i < fitPoints.Count - 1; i++)
                    {
                        double segmentLength = fitPoints[i].DistanceTo(fitPoints[i + 1]);
                        if (currentLength + segmentLength >= distance)
                        {
                            double t = (distance - currentLength) / segmentLength;
                            return new Point3d(
                                fitPoints[i].X + (fitPoints[i + 1].X - fitPoints[i].X) * t,
                                fitPoints[i].Y + (fitPoints[i + 1].Y - fitPoints[i].Y) * t,
                                fitPoints[i].Z + (fitPoints[i + 1].Z - fitPoints[i].Z) * t
                            );
                        }
                        currentLength += segmentLength;
                    }
                    
                    return fitPoints[fitPoints.Count - 1];
                }
                catch
                {
                    return pl.GetPointAtDist(distance);
                }
            }

            /// <summary>
            /// Lấy tiếp tuyến trên polyline tạm tại khoảng cách cho trước
            /// </summary>
            /// <param name="pl">Polyline gốc</param>
            /// <param name="distance">Khoảng cách</param>
            /// <returns>Vector tiếp tuyến</returns>
            private Vector3d GetTangentOnTempPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double distance)
            {
                try
                {
                    // Lấy điểm trước và sau để tính tiếp tuyến
                    double delta = 0.1;
                    Point3d p1 = GetPointOnTempPolyline(pl, Math.Max(0, distance - delta));
                    Point3d p2 = GetPointOnTempPolyline(pl, distance + delta);
                    
                    return p2 - p1;
                }
                catch
                {
                    // Fallback về polyline gốc
                    double param = pl.GetParameterAtPoint(pl.GetPointAtDist(distance));
                            double deltaParam = 0.01;
                            double p0 = Math.Max(pl.StartParam, param - deltaParam);
                            double p1 = Math.Min(pl.EndParam, param + deltaParam);
                    return pl.GetPointAtParameter(p1) - pl.GetPointAtParameter(p0);
                }
            }

            /// <summary>
            /// Tính chiều dài polyline tạm
            /// </summary>
            /// <param name="pl">Polyline gốc</param>
            /// <returns>Chiều dài polyline tạm</returns>
            private double GetTempPolylineLength(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
            {
                try
                {
                    // Tạo danh sách điểm cho polyline fit
                    List<Point3d> fitPoints = new List<Point3d>();
                    
                    // Thêm tất cả các đỉnh của polyline
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(i));
                    }
                    
                    // Nếu polyline đóng, thêm điểm đầu
                    if (pl.Closed && pl.NumberOfVertices > 2)
                    {
                        fitPoints.Add(pl.GetPoint3dAt(0));
                    }
                    
                    // Tính tổng chiều dài polyline fit
                    double totalLength = 0.0;
                    for (int i = 0; i < fitPoints.Count - 1; i++)
                    {
                        totalLength += fitPoints[i].DistanceTo(fitPoints[i + 1]);
                    }
                    
                    return totalLength;
                        }
                        catch
                        {
                    // Fallback về polyline gốc
                    try
                    {
                        return pl.GetDistanceAtParameter(pl.EndParam);
                    }
                    catch
                    {
                        return 0.0;
                    }
                }
            }

            /// <summary>
            /// Tạo danh sách điểm cho polyline fit tạm
            /// </summary>
            /// <param name="pl">Polyline gốc</param>
            /// <returns>Danh sách điểm polyline fit tạm</returns>
            private List<Point3d> CreateTempPolylinePoints(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
            {
                List<Point3d> fitPoints = new List<Point3d>();
                
                try
                {
                    if (pl.NumberOfVertices < 2) return fitPoints;

                    // Tạo danh sách điểm gốc
                    List<Point3d> originalPoints = new List<Point3d>();
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                        originalPoints.Add(pl.GetPoint3dAt(i));
                    }
                    
                    // Nếu polyline đóng, thêm điểm đầu
                    if (pl.Closed && pl.NumberOfVertices > 2)
                    {
                        originalPoints.Add(pl.GetPoint3dAt(0));
                    }
                    
                    // Tạo polyline fit với bán kính bo góc 10%
                    if (originalPoints.Count >= 3)
                    {
                        // Thêm điểm đầu
                        fitPoints.Add(originalPoints[0]);
                        
                        for (int i = 1; i < originalPoints.Count - 1; i++)
                        {
                            Point3d prev = originalPoints[i - 1];
                            Point3d current = originalPoints[i];
                            Point3d next = originalPoints[i + 1];
                            
                            // Tính bán kính bo góc = 10% của đoạn ngắn nhất
                            double dist1 = Math.Sqrt(Math.Pow(prev.X - current.X, 2) + Math.Pow(prev.Y - current.Y, 2) + Math.Pow(prev.Z - current.Z, 2));
                            double dist2 = Math.Sqrt(Math.Pow(current.X - next.X, 2) + Math.Pow(current.Y - next.Y, 2) + Math.Pow(current.Z - next.Z, 2));
                            double radius = Math.Min(dist1, dist2) * 0.1; // 10%
                            
                            // Tạo điểm bo góc
                            List<Point3d> filletPoints = CreateFilletPoints(prev, current, next, radius);
                            
                            if (filletPoints.Count > 0)
                            {
                                fitPoints.AddRange(filletPoints);
                            }
                            else
                            {
                                // Fallback: tạo điểm trung gian đơn giản
                                Point3d midPoint1 = new Point3d(
                                    (prev.X + current.X) / 2,
                                    (prev.Y + current.Y) / 2,
                                    (prev.Z + current.Z) / 2
                                );
                                Point3d midPoint2 = new Point3d(
                                    (current.X + next.X) / 2,
                                    (current.Y + next.Y) / 2,
                                    (current.Z + next.Z) / 2
                                );
                                fitPoints.Add(midPoint1);
                                fitPoints.Add(current);
                                fitPoints.Add(midPoint2);
                            }
                        }
                        
                        // Thêm điểm cuối
                        fitPoints.Add(originalPoints[originalPoints.Count - 1]);
                    }
                    else
                    {
                        // Nếu ít hơn 3 điểm, chỉ copy các điểm gốc
                        for (int i = 0; i < originalPoints.Count; i++)
                        {
                            fitPoints.Add(originalPoints[i]);
                        }
                    }
                }
                catch
                {
                    // Nếu có lỗi, trả về danh sách rỗng
                }
                
                return fitPoints;
            }

            /// <summary>
            /// Tạo điểm bo góc tròn (fillet) cho 3 điểm liên tiếp
            /// </summary>
            /// <param name="prev">Điểm trước</param>
            /// <param name="current">Điểm hiện tại</param>
            /// <param name="next">Điểm tiếp theo</param>
            /// <param name="radius">Bán kính bo góc</param>
            /// <returns>Danh sách điểm bo góc tròn</returns>
            private List<Point3d> CreateFilletPoints(Point3d prev, Point3d current, Point3d next, double radius)
            {
                List<Point3d> filletPoints = new List<Point3d>();
                
                try
                {
                    // Vector từ current đến prev và next
                    Vector3d v1 = (prev - current).GetNormal();
                    Vector3d v2 = (next - current).GetNormal();
                    
                    // Tính góc giữa 2 vector
                    double dotProduct = v1.DotProduct(v2);
                    if (dotProduct > 1.0) dotProduct = 1.0;
                    if (dotProduct < -1.0) dotProduct = -1.0;
                    
                    double angle = Math.Acos(dotProduct);
                    
                    // Kiểm tra góc hợp lệ (nới lỏng điều kiện)
                    if (angle < 0.05 || angle > Math.PI - 0.05)
                    {
                        return filletPoints; // Góc quá nhỏ hoặc quá lớn
                    }
                    
                    // Tính khoảng cách từ current đến điểm bắt đầu bo góc
                    double distance = radius / Math.Tan(angle / 2.0);

                    // Kiểm tra khoảng cách hợp lệ (nới lỏng điều kiện)
                    if (distance <= 0 || distance > radius * 5)
                    {
                        return filletPoints;
                    }
                    
                    // Điểm bắt đầu và kết thúc bo góc
                    Point3d startPoint = current + v1 * distance;
                    Point3d endPoint = current + v2 * distance;
                    
                    // Kiểm tra hướng bo góc (xử lý đúng các góc chỉa lên trên)
                    Vector3d crossProduct = v1.CrossProduct(v2);
                    
                    // Tâm của cung tròn
                    Vector3d bisector = (v1 + v2).GetNormal();
                    double centerDistance = radius / Math.Sin(angle / 2.0);
                    
                    // Đảo ngược hướng bo là chuẩn - không đảo ngược bisector
                    // bisector = -bisector; // Bỏ dòng này
                    
                    Point3d center = current + bisector * centerDistance;
                    
                    // Kiểm tra tâm hợp lệ
                    if (centerDistance <= 0 || centerDistance > radius * 10)
                    {
                        return filletPoints;
                    }
                    
                    // Vector từ tâm đến điểm bắt đầu và kết thúc
                    Vector3d centerToStart = (startPoint - center).GetNormal();
                    Vector3d centerToEnd = (endPoint - center).GetNormal();
                    
                    // Tính góc quay từ start đến end
                    double filletAngle = Math.Acos(centerToStart.DotProduct(centerToEnd));
                    
                    // Kiểm tra hướng quay để đảm bảo cung ra ngoài
                    Vector3d crossCheck = centerToStart.CrossProduct(centerToEnd);
                    bool isClockwise = crossCheck.Z < 0;
                    
                    // Tạo các điểm trên cung tròn thực sự
                    int segments = 16; // Tăng số đoạn để mượt hơn
                    for (int i = 0; i <= segments; i++)
                    {
                        double t = (double)i / segments;
                        double currentAngle = t * filletAngle;
                        
                        // Nếu cung ngược chiều, đảo ngược góc
                        if (isClockwise)
                        {
                            currentAngle = -currentAngle;
                        }
                        
                        // Quay vector centerToStart một góc currentAngle
                        double cosAngle = Math.Cos(currentAngle);
                        double sinAngle = Math.Sin(currentAngle);
                        
                        // Tính vector quay (chỉ trong mặt phẳng XY)
                        Vector3d rotatedVector = new Vector3d(
                            centerToStart.X * cosAngle - centerToStart.Y * sinAngle,
                            centerToStart.X * sinAngle + centerToStart.Y * cosAngle,
                            centerToStart.Z
                        );
                        
                        // Điểm trên cung tròn
                        Point3d point = center + rotatedVector * radius;
                        filletPoints.Add(point);
                    }
                }
                catch
                {
                    // Nếu có lỗi, trả về danh sách rỗng
                }
                
                return filletPoints;
            }

            /// <summary>
            /// Áp vát mềm cho danh sách điểm (giống TAN_SOFT_CHAMFER)
            /// </summary>
            /// <param name="originalPoints">Danh sách điểm gốc</param>
            /// <param name="softness">Bán kính vát mềm</param>
            /// <returns>Danh sách điểm đã vát mềm</returns>
            private List<Point3d> ApplySoftChamferToPoints(List<Point3d> originalPoints, double softness)
            {
                List<Point3d> softPoints = new List<Point3d>();
                
                try
                {
                    if (originalPoints.Count < 3) return originalPoints;
                    
                    // Copy danh sách điểm gốc
                    softPoints.AddRange(originalPoints);
                    
                    // Áp vát mềm toàn bộ (giống TAN_SOFT_CHAMFER)
                    int i = 1;
                    while (i < softPoints.Count - 1)
                    {
                        int iPrev = i - 1;
                        int iMid = i;
                        int iNext = i + 1;
                        int before = softPoints.Count;
                        
                        // Tạo vát mềm tại góc (giống ApplySoftChamfer)
                        List<Point3d> chamferPoints = ApplySoftChamfer(
                            softPoints[iPrev], 
                            softPoints[iMid], 
                            softPoints[iNext], 
                            softness
                        );
                        
                        // Thay thế điểm góc bằng các điểm vát mềm
                        if (chamferPoints.Count > 1)
                        {
                            softPoints.RemoveAt(iMid); // Xóa điểm góc
                            softPoints.InsertRange(iMid, chamferPoints); // Thêm các điểm vát mềm
                        }
                        
                        int after = softPoints.Count;
                        i += (after > before) ? 2 : 1; // Tăng i theo logic TAN_SOFT_CHAMFER
                    }
                }
                catch
                {
                    // Nếu có lỗi, trả về danh sách gốc
                    return originalPoints;
                }
                
                return softPoints;
            }

            /// <summary>
            /// Áp vát mềm tại một góc (giống ApplySoftChamfer)
            /// </summary>
            /// <param name="prev">Điểm trước</param>
            /// <param name="current">Điểm hiện tại</param>
            /// <param name="next">Điểm sau</param>
            /// <param name="softness">Bán kính vát mềm</param>
            /// <returns>Danh sách điểm vát mềm</returns>
            private List<Point3d> ApplySoftChamfer(Point3d prev, Point3d current, Point3d next, double softness)
            {
                List<Point3d> chamferPoints = new List<Point3d>();
                
                try
                {
                    // Vector từ prev đến current và từ current đến next
                    Vector3d v1 = (current - prev).GetNormal();
                    Vector3d v2 = (next - current).GetNormal();

                    // Tính góc giữa 2 vector
                    double angle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, v1.DotProduct(v2))));

                    // Nếu góc quá nhỏ hoặc quá lớn, không tạo vát mềm
                    if (angle < 0.1 || angle > Math.PI - 0.1)
                    {
                        return chamferPoints; // Trả về rỗng để không thay thế điểm
                    }

                    // Tính khoảng cách từ đỉnh đến điểm bắt đầu vát mềm
                    double distance = softness / Math.Tan(angle / 2.0);

                    // Kiểm tra khoảng cách có hợp lệ không
                    if (distance <= 0 || double.IsNaN(distance) || double.IsInfinity(distance))
                    {
                        return chamferPoints; // Trả về rỗng để không thay thế điểm
                    }

                    // Điểm bắt đầu và kết thúc vát mềm
                    Point3d startPoint = current - v1 * distance;
                    Point3d endPoint = current + v2 * distance;

                    // Tâm của cung tròn
                    Vector3d bisector = (v1 + v2).GetNormal();
                    double centerDistance = softness / Math.Sin(angle / 2.0);
                    
                    // Kiểm tra centerDistance có hợp lệ không
                    if (centerDistance <= 0 || double.IsNaN(centerDistance) || double.IsInfinity(centerDistance))
                    {
                        return chamferPoints; // Trả về rỗng để không thay thế điểm
                    }

                    Point3d center = current + bisector * centerDistance;
                    
                    // Tạo các điểm cung tròn
                    Vector3d startVec = (startPoint - center).GetNormal();
                    Vector3d endVec = (endPoint - center).GetNormal();

                    double startAngle = Math.Atan2(startVec.Y, startVec.X);
                    double endAngle = Math.Atan2(endVec.Y, endVec.X);

                    // Đảm bảo góc tăng dần
                    if (endAngle < startAngle) endAngle += 2 * Math.PI;

                    // Tạo các điểm cung tròn - SỐ LƯỢNG ĐIỂM PHỤ THUỘC VÀO GÓC
                    int segments = Math.Max(8, Math.Min(32, (int)(Math.Abs(endAngle - startAngle) * 180 / Math.PI)));
                    double angleStep = (endAngle - startAngle) / segments;

                    for (int i = 0; i <= segments; i++)
                    {
                        double angle1 = startAngle + i * angleStep;
                        Point3d pt = center + new Vector3d(Math.Cos(angle1), Math.Sin(angle1), 0) * softness;
                        chamferPoints.Add(pt);
                    }
                }
                catch
                {
                    // Nếu có lỗi, trả về rỗng để không thay thế điểm
                }

                return chamferPoints;
            }


            /// <summary>
            /// Vẽ polyline tạm từ danh sách điểm
            /// </summary>
            /// <param name="points">Danh sách điểm</param>
            /// <param name="wd">WorldDraw context</param>
            private void DrawTempPolylineFromPoints(List<Point3d> points, WorldDraw wd)
            {
                try
                {
                    if (points.Count < 2) return;

                    // Vẽ polyline tạm màu xanh sáng (giống tick)
                    wd.SubEntityTraits.Color = 4; // Xanh sáng
                    
                    // Vẽ từng đoạn của polyline
                    for (int i = 0; i < points.Count - 1; i++)
                    {
                        wd.Geometry.WorldLine(points[i], points[i + 1]);
                    }
                }
                catch
                {
                    // Bỏ qua lỗi nếu có
                }
            }

            /// <summary>
            /// Tính chiều dài polyline từ danh sách điểm
            /// </summary>
            /// <param name="points">Danh sách điểm</param>
            /// <returns>Chiều dài polyline</returns>
            private double CalculatePolylineLength(List<Point3d> points)
            {
                try
                {
                    double totalLength = 0.0;
                    for (int i = 0; i < points.Count - 1; i++)
                    {
                        totalLength += points[i].DistanceTo(points[i + 1]);
                    }
                    return totalLength;
                }
                catch
                {
                    return 0.0;
                }
            }

            /// <summary>
            /// Lấy điểm trên polyline tại khoảng cách cho trước
            /// </summary>
            /// <param name="points">Danh sách điểm polyline</param>
            /// <param name="distance">Khoảng cách</param>
            /// <returns>Điểm trên polyline</returns>
            private Point3d GetPointOnPolylineAtDistance(List<Point3d> points, double distance)
            {
                try
                {
                    if (points.Count < 2) return points[0];
                    
                    // Tính tổng chiều dài
                    double totalLength = CalculatePolylineLength(points);
                    if (totalLength <= 0) return points[0];
                    
                    // Chuẩn hóa distance
                    if (distance > totalLength) distance = totalLength;
                    
                    // Tìm điểm trên polyline
                    double currentLength = 0.0;
                    for (int i = 0; i < points.Count - 1; i++)
                    {
                        double segmentLength = points[i].DistanceTo(points[i + 1]);
                        if (currentLength + segmentLength >= distance)
                        {
                            double t = (distance - currentLength) / segmentLength;
                            return new Point3d(
                                points[i].X + (points[i + 1].X - points[i].X) * t,
                                points[i].Y + (points[i + 1].Y - points[i].Y) * t,
                                points[i].Z + (points[i + 1].Z - points[i].Z) * t
                            );
                        }
                        currentLength += segmentLength;
                    }
                    
                    return points[points.Count - 1];
                }
                catch
                {
                    return points.Count > 0 ? points[0] : new Point3d(0, 0, 0);
                }
            }

            /// <summary>
            /// Lấy tiếp tuyến trên polyline tại khoảng cách cho trước
            /// </summary>
            /// <param name="points">Danh sách điểm polyline</param>
            /// <param name="distance">Khoảng cách</param>
            /// <returns>Vector tiếp tuyến</returns>
            private Vector3d GetTangentOnPolylineAtDistance(List<Point3d> points, double distance)
            {
                try
                {
                    // Lấy điểm trước và sau để tính tiếp tuyến
                    double delta = 0.1;
                    Point3d p1 = GetPointOnPolylineAtDistance(points, Math.Max(0, distance - delta));
                    Point3d p2 = GetPointOnPolylineAtDistance(points, distance + delta);
                    
                    return p2 - p1;
                }
                catch
                {
                    // Nếu có lỗi, trả về vector mặc định
                    if (points.Count >= 2)
                    {
                        return points[1] - points[0];
                    }
                    return new Vector3d(1, 0, 0);
                }
            }

            /// <summary>
            /// Override method WorldDraw để vẽ thêm các tick vuông góc với polyline (không có fillet arc)
            /// </summary>
            /// <param name="drawable">Đối tượng cần vẽ</param>
            /// <param name="wd">WorldDraw context</param>
            /// <returns>True nếu vẽ thành công</returns>
            public override bool WorldDraw(Drawable drawable, WorldDraw wd)
            {
                // Chỉ áp dụng cho Polyline được phép
                if (drawable is Autodesk.AutoCAD.DatabaseServices.Polyline pl)
                {
                    // Kiểm tra polyline có trong danh sách được phép không
                    if (!_allowedPolylines.Contains(pl.ObjectId))
                    {
                        // Vẽ polyline bình thường nếu không được phép
                        return base.WorldDraw(drawable, wd);
                    }

                    // ẨN polyline chính - KHÔNG gọi base.WorldDraw
                    // Tạo polyline fit tạm một lần
                    List<Point3d> tempPolylinePoints = CreateTempPolylinePoints(pl);
                    
                    // Vẽ polyline tạm màu vàng
                    DrawTempPolylineFromPoints(tempPolylinePoints, wd);
                    
                    // Tính tổng chiều dài polyline tạm
                    double totalLength = CalculatePolylineLength(tempPolylinePoints);
                    if (totalLength <= 0.0) return true;

                    // Khoảng cách giữa các tick (4 inch)
                    double tickSpacing = 4.0;

                    // Số lượng tick
                    int tickCount = (int)Math.Round(totalLength / tickSpacing);
                    if (tickCount < 1) tickCount = 1;
                    double actualSpacing = totalLength / tickCount;

                    // Tham số tick: đỏ 2" + xanh (có thể thay đổi) + đỏ 2"
                    double redLength = 2.0;
                    double blueLength = _blueWidth;
                    double totalTickLength = redLength + blueLength + redLength;
                    double halfTickLength = totalTickLength / 2.0;

                    // Vẽ tick tại các vị trí dọc theo polyline tạm
                    for (int i = 0; i <= tickCount; i++)
                    {
                        double distance = i * actualSpacing;
                        if (distance > totalLength) distance = totalLength;

                        try
                        {
                            // Lấy điểm trên polyline tạm tại khoảng cách distance
                            Point3d pointOnTempPl = GetPointOnPolylineAtDistance(tempPolylinePoints, distance);

                            // Tính hướng tiếp tuyến tại điểm này trên polyline tạm
                            Vector3d tangent = GetTangentOnPolylineAtDistance(tempPolylinePoints, distance);
                            if (tangent.Length < 1e-9) continue;

                            tangent = tangent.GetNormal();

                            // Vector pháp tuyến (vuông góc với tiếp tuyến)
                            Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0);

                            // Tính các điểm của tick
                            Point3d tickStart = pointOnTempPl - normal * halfTickLength;
                            Point3d redEnd1 = pointOnTempPl - normal * (halfTickLength - redLength);
                            Point3d blueStart = pointOnTempPl - normal * (halfTickLength - redLength);
                            Point3d blueEnd = pointOnTempPl + normal * (halfTickLength - redLength);
                            Point3d redStart2 = pointOnTempPl + normal * (halfTickLength - redLength);
                            Point3d tickEnd = pointOnTempPl + normal * halfTickLength;

                            // Vẽ đoạn đỏ đầu tiên
                            wd.SubEntityTraits.Color = 1; // Đỏ
                            wd.Geometry.WorldLine(tickStart, redEnd1);

                            // Vẽ đoạn xanh giữa
                            wd.SubEntityTraits.Color = 4; // Xanh sáng
                            wd.Geometry.WorldLine(blueStart, blueEnd);

                            // Vẽ đoạn đỏ cuối
                            wd.SubEntityTraits.Color = 1; // Đỏ
                            wd.Geometry.WorldLine(redStart2, tickEnd);
                        }
                        catch
                        {
                            // Bỏ qua lỗi nếu có
                            continue;
                        }
                    }
                }
                return true;
            }


            /// <summary>
            /// Vẽ spline ảo theo polyline (làm mượt các góc) - Giữ lại cho tương thích
            /// </summary>
            /// <param name="pl">Polyline cần vẽ spline</param>
            /// <param name="wd">WorldDraw context</param>
            private void DrawVirtualSpline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, WorldDraw wd)
            {
                if (pl.NumberOfVertices < 3) return;

                // Tạo danh sách điểm cho spline
                Point3dCollection splinePoints = new Point3dCollection();

                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    Point3d pt = pl.GetPoint3dAt(i);
                    splinePoints.Add(pt);
                }

                // Vẽ spline ảo bằng cách vẽ nhiều đoạn thẳng mượt
                int segments = 20; // Số đoạn để tạo spline mượt
                for (int i = 0; i < splinePoints.Count - 1; i++)
                {
                    Point3d p1 = splinePoints[i];
                    Point3d p2 = splinePoints[i + 1];

                    // Vẽ đoạn thẳng với các điểm trung gian để tạo hiệu ứng mượt
                    for (int j = 0; j < segments; j++)
                    {
                        double t1 = (double)j / segments;
                        double t2 = (double)(j + 1) / segments;

                        Point3d pt1 = p1 + (p2 - p1) * t1;
                        Point3d pt2 = p1 + (p2 - p1) * t2;

                        // Thêm độ cong nhẹ để tạo hiệu ứng spline
                        if (i > 0 && i < splinePoints.Count - 2)
                        {
                            Vector3d normal = new Vector3d(-(p2 - p1).Y, (p2 - p1).X, 0.0).GetNormal();
                            double curve1 = Math.Sin(t1 * Math.PI) * 0.5;
                            double curve2 = Math.Sin(t2 * Math.PI) * 0.5;

                            pt1 = pt1 + normal * curve1;
                            pt2 = pt2 + normal * curve2;
                        }

                        wd.Geometry.WorldLine(pt1, pt2);
                    }
                }
            }

            /// <summary>
            /// Thiết lập bộ lọc tùy chỉnh (có thể thêm logic lọc nâng cao)
            /// </summary>
            private new void SetCustomFilter()
            {
                // Có thể thêm logic lọc nâng cao (ví dụ layer cụ thể) nếu muốn
            }
        }

        /// <summary>
        /// Class chứa các lệnh để bật/tắt FilletOverrule
        /// </summary>
        public class FilletOverruleCmd
        {
            private static FilletOverrule _overrule;
            private static double _blueWidth = 8.0; // Độ dài đoạn xanh mặc định


            /// <summary>
            /// Lệnh vẽ duct với tick: nhập W -> chọn điểm đầu tiên -> vẽ polyline với tick overrule
            /// </summary>
            [CommandMethod("TANDUCTGOOD")]
            public static void TanDuctGood()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Bước 1: Nhập W (width đoạn xanh)
                    PromptDoubleOptions widthOpts = new PromptDoubleOptions("\nNhập W (độ dài đoạn xanh inch): ");
                    widthOpts.AllowNegative = false;
                    widthOpts.AllowZero = false;
                    widthOpts.DefaultValue = 8.0;

                    PromptDoubleResult widthResult = ed.GetDouble(widthOpts);
                    if (widthResult.Status != PromptStatus.OK) return;

                    _blueWidth = widthResult.Value;
                    ed.WriteMessage($"\nĐã đặt W = {_blueWidth:F1}\"");

                    // Bước 2: Chọn điểm đầu tiên
                    PromptPointOptions firstPointOpts = new PromptPointOptions("\nChọn điểm đầu tiên: ");
                    firstPointOpts.AllowNone = false;

                    PromptPointResult firstPointResult = ed.GetPoint(firstPointOpts);
                    if (firstPointResult.Status != PromptStatus.OK) return;

                    Point3d firstPoint = firstPointResult.Value;
                    ed.WriteMessage($"\nĐã chọn điểm đầu tiên: ({firstPoint.X:F2}, {firstPoint.Y:F2})");

                    // Bước 3: Bật overrule
                    if (_overrule == null)
                    {
                        _overrule = new FilletOverrule();
                        _overrule.UpdateBlueWidth(_blueWidth);
                        Overrule.AddOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule, true);
                        Overrule.Overruling = true;
                        ed.WriteMessage($"\nTick overrule đã bật W={_blueWidth:F1}\"");
                    }
                    else
                    {
                        _overrule.UpdateBlueWidth(_blueWidth);
                        ed.WriteMessage($"\nĐã cập nhật W = {_blueWidth:F1}\"");
                    }

                    // Bước 4: Vẽ polyline với DrawJig
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // Tạo polyline
                        Autodesk.AutoCAD.DatabaseServices.Polyline pl = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                        pl.SetDatabaseDefaults();
                        pl.Color = Color.FromColorIndex(ColorMethod.ByAci, 4); // Màu xanh sáng

                        // Thêm điểm đầu tiên
                        pl.AddVertexAt(0, new Point2d(firstPoint.X, firstPoint.Y), 0.0, 0.0, 0.0);

                        // Sử dụng DrawJig để vẽ polyline tương tác
                        var jig = new DuctDrawJig(pl, _blueWidth);

                        // Vòng lặp vẽ polyline
                        bool finished = false;
                        while (!finished)
                        {
                            PromptResult pr = ed.Drag(jig);
                            switch (pr.Status)
                            {
                                case PromptStatus.OK: // User click để chọn điểm
                                    jig.AcceptPoint();
                                    break;
                                case PromptStatus.None: // Space/Enter để kết thúc
                                    finished = true;
                                    break;
                                case PromptStatus.Cancel: // ESC để hủy
                                    pl.Dispose();
                                    tr.Abort();
                                    return;
                                default:
                                    finished = true;
                                    break;
                            }
                        }

                        // Kiểm tra polyline có ít nhất 2 điểm không
                        if (pl.NumberOfVertices >= 2)
                        {
                            // Xóa điểm preview cuối cùng nếu có
                            if (jig.HasPreviewPoint)
                            {
                                pl.RemoveVertexAt(pl.NumberOfVertices - 1);
                            }

                            // Thêm polyline vào ModelSpace
                            btr.AppendEntity(pl);
                            tr.AddNewlyCreatedDBObject(pl, true);

                            // Thêm polyline vào danh sách được phép hiển thị tick
                            FilletOverrule.AddAllowedPolyline(pl.ObjectId);

                            tr.Commit();
                            ed.WriteMessage($"\nĐã tạo duct với {pl.NumberOfVertices} điểm! Tick sẽ hiển thị tự động.");
                        }
                        else
                        {
                            pl.Dispose();
                            tr.Abort();
                            ed.WriteMessage("\nCần ít nhất 2 điểm để tạo duct!");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// DrawJig để vẽ polyline tương tác cho duct
        /// </summary>
        public class DuctDrawJig : Autodesk.AutoCAD.EditorInput.DrawJig
        {
            private Autodesk.AutoCAD.DatabaseServices.Polyline _polyline;
            private double _blueWidth;
            private Point3d _currentPoint;
            private bool _hasPreviewPoint = false;

            public bool HasPreviewPoint { get { return _hasPreviewPoint; } }

            public DuctDrawJig(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double blueWidth)
            {
                _polyline = pl;
                _blueWidth = blueWidth;
            }

            public void AcceptPoint()
            {
                if (_polyline.NumberOfVertices == 0) return;

                // Cố định điểm preview hiện tại
                Point2d lastPoint = _polyline.GetPoint2dAt(_polyline.NumberOfVertices - 1);
                _polyline.AddVertexAt(_polyline.NumberOfVertices, lastPoint, 0.0, 0.0, 0.0);
                _hasPreviewPoint = true;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                JigPromptPointOptions jppo = new JigPromptPointOptions();

                if (_polyline.NumberOfVertices == 1)
                {
                    jppo.Message = "\nChọn điểm tiếp theo (Space/Enter để kết thúc): ";
                }
                else
                {
                    jppo.Message = "\nChọn điểm tiếp theo (Space/Enter để kết thúc): ";
                }

                jppo.UserInputControls = UserInputControls.NullResponseAccepted;

                if (_polyline.NumberOfVertices > 0)
                {
                    jppo.UseBasePoint = true;
                    jppo.BasePoint = _polyline.GetPoint3dAt(_polyline.NumberOfVertices - 1);
                }

                PromptPointResult ppr = prompts.AcquirePoint(jppo);

                if (ppr.Status == PromptStatus.OK)
                {
                    Point3d newPoint = ppr.Value;
                    if (newPoint == _currentPoint)
                        return SamplerStatus.NoChange;

                    _currentPoint = newPoint;

                    // Cập nhật polyline
                    if (_polyline.NumberOfVertices == 1)
                    {
                        // Thêm điểm preview
                        _polyline.AddVertexAt(1, new Point2d(_currentPoint.X, _currentPoint.Y), 0.0, 0.0, 0.0);
                        _hasPreviewPoint = true;
                    }
                    else if (_polyline.NumberOfVertices > 1)
                    {
                        // Cập nhật điểm preview cuối cùng
                        int lastIndex = _polyline.NumberOfVertices - 1;
                        _polyline.SetPointAt(lastIndex, new Point2d(_currentPoint.X, _currentPoint.Y));
                    }

                    return SamplerStatus.OK;
                }
                else if (ppr.Status == PromptStatus.None)
                {
                    return SamplerStatus.Cancel;
                }

                return SamplerStatus.Cancel;
            }


            protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw wd)
            {
                // Vẽ polyline chính (màu trắng)
                if (_polyline.NumberOfVertices >= 2)
                {
                    for (int i = 1; i < _polyline.NumberOfVertices; i++)
                    {
                        Point3d p1 = _polyline.GetPoint3dAt(i - 1);
                        Point3d p2 = _polyline.GetPoint3dAt(i);
                        wd.Geometry.WorldLine(p1, p2);
                    }
                }

                // Vẽ polyline tạm với fillet (màu vàng) - chỉ preview
                if (_polyline.NumberOfVertices >= 2)
                {
                    DrawFilletPolylinePreview(wd);
                }

                // Vẽ tick preview
                if (_polyline.NumberOfVertices >= 2)
                {
                    DrawTickPreview(wd);
                }

                return true;
            }

            /// <summary>
            /// Vẽ preview polyline fillet màu vàng
            /// </summary>
            private void DrawFilletPolylinePreview(Autodesk.AutoCAD.GraphicsInterface.WorldDraw wd)
            {
                try
                {
                    // Tạo danh sách điểm từ polyline hiện tại
                    List<Point3d> points = new List<Point3d>();
                    for (int i = 0; i < _polyline.NumberOfVertices; i++)
                    {
                        points.Add(_polyline.GetPoint3dAt(i));
                    }

                    // Tạo polyline fillet preview
                    if (points.Count >= 2)
                    {
                        // Vẽ polyline fillet màu vàng - vẽ đúng vị trí với polyline gốc
                        wd.SubEntityTraits.Color = 2; // Vàng
                        
                        // Vẽ từng đoạn của polyline gốc
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                            wd.Geometry.WorldLine(points[i], points[i + 1]);
                        }
                        
                        // Vẽ đoạn cuối nếu có điểm preview
                        if (_hasPreviewPoint && _polyline.NumberOfVertices > 1)
                        {
                            Point3d lastPoint = _polyline.GetPoint3dAt(_polyline.NumberOfVertices - 1);
                            wd.Geometry.WorldLine(lastPoint, _currentPoint);
                        }
                    }
                }
                catch
                {
                    // Bỏ qua lỗi nếu có
                }
            }




            private void DrawTickPreview(Autodesk.AutoCAD.GraphicsInterface.WorldDraw wd)
            {
                // Tính tổng chiều dài polyline hiện tại
                double totalLength = 0.0;
                for (int i = 1; i < _polyline.NumberOfVertices; i++)
                {
                    Point3d p1 = _polyline.GetPoint3dAt(i - 1);
                    Point3d p2 = _polyline.GetPoint3dAt(i);
                    totalLength += (p2 - p1).Length;
                }

                if (totalLength <= 0.0) return;

                // Khoảng cách giữa các tick (4 inch)
                double tickSpacing = 4.0;
                int tickCount = (int)Math.Round(totalLength / tickSpacing);
                if (tickCount < 1) tickCount = 1;
                double actualSpacing = totalLength / tickCount;

                // Tham số tick
                double redLength = 2.0;
                double blueLength = _blueWidth;
                double totalTickLength = redLength + blueLength + redLength;
                double halfTickLength = totalTickLength / 2.0;

                // Vẽ tick preview
                for (int i = 0; i <= tickCount; i++)
                {
                    double distance = i * actualSpacing;
                    if (distance > totalLength) distance = totalLength;

                    try
                    {
                        // Tìm điểm trên polyline tại khoảng cách distance
                        double walked = 0.0;
                        Point3d pointOnPl = _polyline.GetPoint3dAt(0);

                        for (int j = 1; j < _polyline.NumberOfVertices; j++)
                        {
                            Point3d p1 = _polyline.GetPoint3dAt(j - 1);
                            Point3d p2 = _polyline.GetPoint3dAt(j);
                            double segLength = (p2 - p1).Length;

                            if (walked + segLength >= distance)
                            {
                                double t = (distance - walked) / segLength;
                                pointOnPl = p1 + (p2 - p1) * t;
                                break;
                            }
                            walked += segLength;
                        }

                        // Tính hướng tiếp tuyến tại điểm hiện tại
                        Vector3d tangent = new Vector3d(1, 0, 0); // Mặc định

                        // Tìm đoạn chứa điểm hiện tại
                        double walked2 = 0.0;
                        for (int k = 1; k < _polyline.NumberOfVertices; k++)
                        {
                            Point3d p1 = _polyline.GetPoint3dAt(k - 1);
                            Point3d p2 = _polyline.GetPoint3dAt(k);
                            double segLength = (p2 - p1).Length;

                            if (walked2 + segLength >= distance)
                            {
                                // Điểm nằm trong đoạn này
                                tangent = (p2 - p1).GetNormal();
                                break;
                            }
                            walked2 += segLength;
                        }

                        // Vector pháp tuyến
                        Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0);

                        // Tính các điểm của tick
                        Point3d tickStart = pointOnPl - normal * halfTickLength;
                        Point3d redEnd1 = pointOnPl - normal * (halfTickLength - redLength);
                        Point3d blueStart = pointOnPl - normal * (halfTickLength - redLength);
                        Point3d blueEnd = pointOnPl + normal * (halfTickLength - redLength);
                        Point3d redStart2 = pointOnPl + normal * (halfTickLength - redLength);
                        Point3d tickEnd = pointOnPl + normal * halfTickLength;

                        // Vẽ tick
                        wd.SubEntityTraits.Color = 1; // Đỏ
                        wd.Geometry.WorldLine(tickStart, redEnd1);
                        wd.Geometry.WorldLine(redStart2, tickEnd);

                        wd.SubEntityTraits.Color = 4; // Xanh sáng
                        wd.Geometry.WorldLine(blueStart, blueEnd);
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
    }
}