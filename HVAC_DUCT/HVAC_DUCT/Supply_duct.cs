using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TAN2025_HVAC_DUCT_SUPPLY_AIR
{
    /// <summary>
    /// Main application class for HVAC Duct system
    /// </summary>
    public class Supply_duct : IExtensionApplication
{
        public void Initialize()
        {
            try
            {
                // Khởi tạo overrule
                FilletOverruleCmd.InitializeOverrule();
                
                // Auto load tick khi load DLL
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    var ed = doc.Editor;
                    ed.WriteMessage("\nAuto loading HVAC_DUCT ticks...");
                    
                    // Gọi trực tiếp method thay vì command
                    FilletOverruleCmd.LoadTempTicks();
                    
                    ed.WriteMessage("\nAuto load completed!");
                }
            }
            catch (System.Exception ex)
            {
                HandleException("Initialize", ex);
            }
        }

        public void Terminate()
        {
            try
            {
                // Cleanup overrule
                FilletOverruleCmd.CleanupOverrule();
            }
            catch (System.Exception ex)
            {
                HandleException("Terminate", ex);
            }
        }

        /// <summary>
        /// Centralized exception handling
        /// </summary>
        /// <param name="methodName">Name of the method where exception occurred</param>
        /// <param name="ex">Exception object</param>
        private static void HandleException(string methodName, System.Exception ex)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    var ed = doc.Editor;
                    ed.WriteMessage($"\nError in {methodName}: {ex.Message}");
                }
                System.Diagnostics.Debug.WriteLine($"Error in {methodName}: {ex.Message}");
            }
            catch
            {
                // Fallback error handling
                System.Diagnostics.Debug.WriteLine($"Critical error in {methodName}: {ex.Message}");
            }
        }
    /// <summary>
    /// DrawableOverrule để vẽ thêm các tick vuông góc với Polyline (không có fillet arc)
    /// </summary>
    public class FilletOverrule : DrawableOverrule
    {
        #region Constants
        private const double DEFAULT_BLUE_WIDTH = 8.0;
        private const double TICK_SPACING = 4.0;
        private const double RED_LENGTH = 2.0;
        private const int FILLET_SEGMENTS = 16;
        private const double FILLET_RADIUS_PERCENT = 0.1;
        #endregion

        #region Private Fields
        private static double _blueWidth = DEFAULT_BLUE_WIDTH;
        private static readonly HashSet<ObjectId> _allowedPolylines = new HashSet<ObjectId>();
        private static readonly Dictionary<ObjectId, double> _polylineWidths = new Dictionary<ObjectId, double>();
        #endregion

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
            public static void UpdateBlueWidth(double newWidth)
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
            /// Thêm polyline với width riêng vào danh sách được phép hiển thị tick
        /// </summary>
            /// <param name="polylineId">ObjectId của polyline</param>
            /// <param name="width">Width riêng cho polyline này</param>
            public static void AddAllowedPolylineWithWidth(ObjectId polylineId, double width)
            {
                _allowedPolylines.Add(polylineId);
                _polylineWidths[polylineId] = width;
            }

        /// <summary>
        /// Xóa polyline khỏi danh sách hiển thị tick
        /// </summary>
        /// <param name="polylineId">ObjectId của polyline</param>
        public static void RemoveAllowedPolyline(ObjectId polylineId)
        {
            _allowedPolylines.Remove(polylineId);
            _polylineWidths.Remove(polylineId);
        }

        /// <summary>
            /// Lấy width riêng của polyline
        /// </summary>
            /// <param name="polylineId">ObjectId của polyline</param>
            /// <returns>Width của polyline hoặc width mặc định</returns>
            public static double GetPolylineWidth(ObjectId polylineId)
        {
                return _polylineWidths.ContainsKey(polylineId) ? _polylineWidths[polylineId] : _blueWidth;
        }

        /// <summary>
            /// Xóa tất cả polyline khỏi danh sách
        /// </summary>
        public static void ClearAllowedPolylines()
        {
            _allowedPolylines.Clear();
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
                var fitPoints = CreateTempPolylinePoints(pl);
                return GetPointOnPolylineAtDistance(fitPoints, distance);
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
                var fitPoints = CreateTempPolylinePoints(pl);
                return GetTangentOnPolylineAtDistance(fitPoints, distance);
            }
            catch
            {
                // Fallback về polyline gốc
                var param = pl.GetParameterAtPoint(pl.GetPointAtDist(distance));
                var deltaParam = 0.01;
                var p0 = Math.Max(pl.StartParam, param - deltaParam);
                var p1 = Math.Min(pl.EndParam, param + deltaParam);
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
                var fitPoints = CreateTempPolylinePoints(pl);
                return CalculatePolylineLength(fitPoints);
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
        public List<Point3d> CreateTempPolylinePoints(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            var fitPoints = new List<Point3d>();
            
            try
            {
                if (pl.NumberOfVertices < 2) return fitPoints;

                // Tạo danh sách điểm gốc
                var originalPoints = new List<Point3d>();
                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    originalPoints.Add(pl.GetPoint3dAt(i));
                }
                
                // Nếu polyline đóng, thêm điểm đầu
                if (pl.Closed && pl.NumberOfVertices > 2)
                {
                    originalPoints.Add(pl.GetPoint3dAt(0));
                }
                
                // Tạo polyline fit với bán kính bo góc
                if (originalPoints.Count >= 3)
                {
                    // Thêm điểm đầu
                    fitPoints.Add(originalPoints[0]);
                    
                    for (int i = 1; i < originalPoints.Count - 1; i++)
                    {
                        var prev = originalPoints[i - 1];
                        var current = originalPoints[i];
                        var next = originalPoints[i + 1];
                        
                        // Tính bán kính bo góc = FILLET_RADIUS_PERCENT của đoạn ngắn nhất
                        var dist1 = prev.DistanceTo(current);
                        var dist2 = current.DistanceTo(next);
                        var radius = Math.Min(dist1, dist2) * FILLET_RADIUS_PERCENT;
                        
                        // Tạo điểm bo góc
                        var filletPoints = CreateFilletPoints(prev, current, next, radius);
                        
                        if (filletPoints.Count > 0)
                        {
                            fitPoints.AddRange(filletPoints);
                        }
                        else
                        {
                            // Fallback: tạo điểm trung gian đơn giản
                            var midPoint1 = new Point3d(
                                (prev.X + current.X) / 2,
                                (prev.Y + current.Y) / 2,
                                (prev.Z + current.Z) / 2
                            );
                            var midPoint2 = new Point3d(
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
                    fitPoints.AddRange(originalPoints);
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
                    for (int i = 0; i <= FILLET_SEGMENTS; i++)
                    {
                        var t = (double)i / FILLET_SEGMENTS;
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

                    // Số lượng tick
                    var tickCount = (int)Math.Round(totalLength / TICK_SPACING);
                    if (tickCount < 1) tickCount = 1;
                    var actualSpacing = totalLength / tickCount;

                    // Tham số tick: đỏ + xanh + đỏ
                    var blueLength = GetPolylineWidth(pl.ObjectId);
                    var totalTickLength = RED_LENGTH + blueLength + RED_LENGTH;
                    var halfTickLength = totalTickLength / 2.0;

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
                            var tickStart = pointOnTempPl - normal * halfTickLength;
                            var redEnd1 = pointOnTempPl - normal * (halfTickLength - RED_LENGTH);
                            var blueStart = pointOnTempPl - normal * (halfTickLength - RED_LENGTH);
                            var blueEnd = pointOnTempPl + normal * (halfTickLength - RED_LENGTH);
                            var redStart2 = pointOnTempPl + normal * (halfTickLength - RED_LENGTH);
                            var tickEnd = pointOnTempPl + normal * halfTickLength;

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
            #region Constants
            private const string APP_NAME = "HVAC_DUCT_SUPPLY_AIR";
            private const double DEFAULT_BLUE_WIDTH = 8.0;
            #endregion

            #region Private Fields
            private static FilletOverrule _overrule;
            private static double _blueWidth = DEFAULT_BLUE_WIDTH;
            #endregion

        /// <summary>
        /// Khởi tạo overrule
        /// </summary>
        public static void InitializeOverrule()
        {
            try
            {
                if (_overrule == null)
                {
                    _overrule = new FilletOverrule();
                    Overrule.AddOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule, true);
                    Overrule.AddOverrule(RXClass.GetClass(typeof(Polyline2d)), _overrule, true);
                    _overrule.SetCustomFilter();
                }
            }
            catch (System.Exception ex)
            {
                Supply_duct.HandleException("InitializeOverrule", ex);
            }
        }

        /// <summary>
        /// Cleanup overrule
        /// </summary>
        public static void CleanupOverrule()
        {
            try
            {
                if (_overrule != null)
                {
                    Overrule.RemoveOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule);
                    Overrule.RemoveOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline2d)), _overrule);
                    _overrule = null;
                }
            }
            catch (System.Exception ex)
            {
                Supply_duct.HandleException("CleanupOverrule", ex);
            }
        }

        /// <summary>
        /// Đăng ký ứng dụng trong database
        /// </summary>
        /// <param name="db">Database</param>
        private static void RegisterApplication(Database db)
        {
            try
            {
                using (var regTr = db.TransactionManager.StartTransaction())
                {
                    var regAppTable = regTr.GetObject(db.RegAppTableId, OpenMode.ForRead) as RegAppTable;
                    if (!regAppTable.Has(APP_NAME))
                    {
                        regAppTable.UpgradeOpen();
                        var regAppRecord = new RegAppTableRecord();
                        regAppRecord.Name = APP_NAME;
                        regAppTable.Add(regAppRecord);
                        regTr.AddNewlyCreatedDBObject(regAppRecord, true);
                    }
                    regTr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                Supply_duct.HandleException("RegisterApplication", ex);
            }
        }

        /// <summary>
        /// Lưu thông tin tick và MText ObjectId vào XData của polyline
        /// </summary>
        /// <param name="pl">Polyline cần lưu</param>
        /// <param name="blueWidth">Width của tick</param>
        /// <param name="mtextId">ObjectId của MText (có thể null)</param>
        private static void SaveTickInfoToDatabase(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double blueWidth, ObjectId? mtextId = null)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                
                // Đăng ký ứng dụng trước khi lưu XData
                string appName = "HVAC_DUCT_SUPPLY_AIR";
                using (Transaction regTr = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regAppTable = regTr.GetObject(db.RegAppTableId, OpenMode.ForRead) as RegAppTable;
                    if (!regAppTable.Has(appName))
                    {
                        regAppTable.UpgradeOpen();
                        RegAppTableRecord regAppRecord = new RegAppTableRecord();
                        regAppRecord.Name = appName;
                        regAppTable.Add(regAppRecord);
                        regTr.AddNewlyCreatedDBObject(regAppRecord, true);
                    }
                    regTr.Commit();
                }
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Mở polyline để ghi XData
                    var plForWrite = tr.GetObject(pl.ObjectId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    
        // Tạo ResultBuffer với tên ứng dụng, width và MText ObjectId
        ResultBuffer rb = new ResultBuffer();
        rb.Add(new TypedValue(1001, APP_NAME)); // Tên ứng dụng
        rb.Add(new TypedValue(1040, blueWidth)); // Width as Double
                    
                    if (mtextId.HasValue && mtextId.Value.IsValid)
                    {
                        rb.Add(new TypedValue(1005, mtextId.Value.Handle.Value)); // MText Handle
                    }
                    
                    // Thêm XData vào polyline
                    plForWrite.XData = rb;
                    
                    tr.Commit();
                    
                    // Debug message
                    Document doc2 = Application.DocumentManager.MdiActiveDocument;
                    Editor ed2 = doc2.Editor;
                    if (mtextId.HasValue)
                    {
                        ed2.WriteMessage($"\nĐã lưu width {blueWidth:F1} và MText Handle {mtextId.Value.Handle.Value} vào XData của Handle: {pl.ObjectId.Handle.Value}");
                    }
                    else
                    {
                        ed2.WriteMessage($"\nĐã lưu width {blueWidth:F1} vào XData của Handle: {pl.ObjectId.Handle.Value}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Log lỗi để debug
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi lưu XData: {ex.Message}");
                }
            }


            /// <summary>
            /// Lệnh vẽ duct với tick: nhập W -> chọn điểm đầu tiên -> vẽ polyline với tick overrule
            /// </summary>
            [CommandMethod("TAN25_HVAC_DUCT_SUPPLY_AIR")]
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
                        FilletOverrule.UpdateBlueWidth(_blueWidth);
                        Overrule.AddOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule, true);
                        Overrule.Overruling = true;
                        ed.WriteMessage($"\nTick overrule đã bật W={_blueWidth:F1}\"");
                    }
                    else
                    {
                        FilletOverrule.UpdateBlueWidth(_blueWidth);
                        ed.WriteMessage($"\nĐã cập nhật W = {_blueWidth:F1}\"");
                    }

                    // Bước 4: Vẽ polyline với DrawJig
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // Tạo polyline
                        var pl = new Autodesk.AutoCAD.DatabaseServices.Polyline();
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

                            // Thêm polyline vào danh sách được phép hiển thị tick với width riêng
                            FilletOverrule.AddAllowedPolylineWithWidth(pl.ObjectId, _blueWidth);
                            
                            // Lưu thông tin tick vào XData
                            SaveTickInfoToDatabase(pl, _blueWidth);
                    
                    tr.Commit();
                            
                            // Tạo MText sau khi kết thúc lệnh vẽ polyline
                            CreateMTextForPolyline(pl);
                            
                            ed.WriteMessage($"\nĐã tạo duct với {pl.NumberOfVertices} điểm! Tick sẽ hiển thị tự động.");
                            
                            // Refresh màn hình để hiển thị tick và MText
                            ed.Regen();
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


        /// <summary>
        /// Cập nhật width trong XData của polyline
        /// </summary>
        /// <param name="pl">Polyline cần cập nhật</param>
        /// <param name="newWidth">Width mới</param>
        private static void UpdateWidthInXData(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double newWidth)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                // Đăng ký ứng dụng nếu chưa có
                string appName = "HVAC_DUCT_SUPPLY_AIR";
                using (Transaction regTr = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regAppTable = regTr.GetObject(db.RegAppTableId, OpenMode.ForRead) as RegAppTable;
                    if (!regAppTable.Has(appName))
                    {
                        regAppTable.UpgradeOpen();
                        RegAppTableRecord regAppRecord = new RegAppTableRecord();
                        regAppRecord.Name = appName;
                        regAppTable.Add(regAppRecord);
                        regTr.AddNewlyCreatedDBObject(regAppRecord, true);
                    }
                    regTr.Commit();
                }

        // Lấy MText ObjectId hiện có từ XData
        ObjectId? existingMTextId = LoadMTextIdFromXData(pl);

        // Tạo ResultBuffer với tên ứng dụng, width mới và MText ObjectId (nếu có)
        ResultBuffer rb = new ResultBuffer();
        rb.Add(new TypedValue(1001, "HVAC_DUCT_SUPPLY_AIR")); // Tên ứng dụng
        rb.Add(new TypedValue(1040, newWidth)); // Width as Double
                
                if (existingMTextId.HasValue && existingMTextId.Value.IsValid)
                {
                    rb.Add(new TypedValue(1005, existingMTextId.Value.Handle.Value)); // MText Handle
                }

                // Cập nhật XData
                pl.XData = rb;
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi cập nhật XData: {ex.Message}");
            }
        }

        /// <summary>
        /// Xóa MText cũ của polyline
        /// </summary>
        /// <param name="pl">Polyline cần xóa MText</param>
        private static void DeleteMTextForPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Tìm và xóa MText gần polyline
                    Point3d endPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                    Vector3d tangent = GetTangentAtEnd(pl);
                    Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0).GetNormal();
                    Point3d expectedTextPosition = endPoint + normal * 5.0;

                    List<ObjectId> mtextToDelete = new List<ObjectId>();

                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass.Name == "AcDbMText")
                        {
                            MText mtext = tr.GetObject(objId, OpenMode.ForRead) as MText;
                            if (mtext != null)
                            {
                                // Kiểm tra khoảng cách (trong vòng 2 inch)
                                double distance = mtext.Location.DistanceTo(expectedTextPosition);
                                if (distance < 2.0)
                                {
                                    mtextToDelete.Add(objId);
                                }
                            }
                        }
                    }

                    // Xóa MText
                    foreach (ObjectId objId in mtextToDelete)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                        entity.Erase();
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi xóa MText: {ex.Message}");
            }
        }

        /// <summary>
        /// Load MText ObjectId từ XData của polyline
        /// </summary>
        /// <param name="pl">Polyline cần load MText ObjectId</param>
        /// <returns>MText ObjectId hoặc null nếu không tìm thấy</returns>
        private static ObjectId? LoadMTextIdFromXData(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            try
            {
                if (pl.XData == null) return null;

                ResultBuffer rb = pl.XData;
                bool isHVACDuct = false;
                ObjectId? mtextId = null;

                foreach (var tv in rb)
                {
                    if (tv.TypeCode == 1001) // Tên ứng dụng
                    {
                        if (tv.Value.ToString() == APP_NAME)
                        {
                            isHVACDuct = true;
                        }
                    }
                    else if (tv.TypeCode == 1005) // Handle - MText ObjectId
                    {
                        try
                        {
                            long handleValue = Convert.ToInt64(tv.Value);
                            mtextId = pl.Database.GetObjectId(false, new Handle(handleValue), 0);
                        }
                        catch
                        {
                            // Handle không hợp lệ
                        }
                    }
                }

                // Chỉ trả về MText ObjectId nếu là ứng dụng HVAC_DUCT
                return isHVACDuct ? mtextId : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Load width từ XData của polyline
        /// </summary>
        /// <param name="pl">Polyline cần load width</param>
        /// <returns>Width từ XData hoặc 0 nếu không tìm thấy</returns>
        private static double LoadWidthFromXData(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            try
            {
                if (pl.XData == null) return 0;

                ResultBuffer rb = pl.XData;
                bool isHVACDuct = false;
                double width = 0;

                foreach (var tv in rb)
                {
                    if (tv.TypeCode == 1001) // Tên ứng dụng
                    {
                        if (tv.Value.ToString() == APP_NAME)
                        {
                            isHVACDuct = true;
                        }
                    }
                    else if (tv.TypeCode == 1040) // Double - Width
                    {
                        width = (double)tv.Value;
                    }
                }

                // Chỉ trả về width nếu là ứng dụng HVAC_DUCT
                return isHVACDuct ? width : 0;
            }
            catch
            {
                return 0;
            }
        }


        /// <summary>
        /// Tạo hoặc lấy layer ID
        /// </summary>
        /// <param name="layerName">Tên layer</param>
        /// <param name="colorIndex">Màu layer</param>
        /// <returns>ObjectId của layer</returns>
        private static ObjectId CreateOrGetLayer(string layerName, short colorIndex)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                
                if (layerTable.Has(layerName))
                {
                    return layerTable[layerName];
                }
                else
                {
                    layerTable.UpgradeOpen();
                    LayerTableRecord layerRecord = new LayerTableRecord();
                    layerRecord.Name = layerName;
                    layerRecord.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                    layerTable.Add(layerRecord);
                    tr.AddNewlyCreatedDBObject(layerRecord, true);
                    tr.Commit();
                    return layerRecord.ObjectId;
                }
            }
        }

        /// <summary>
        /// Cập nhật MText trực tiếp bằng ObjectId từ XData
        /// </summary>
        /// <param name="pl">Polyline cần cập nhật MText</param>
        /// <param name="newWidth">Width mới</param>
        private static void UpdateMTextForPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double newWidth)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                // Lấy MText ObjectId từ XData
                ObjectId? mtextId = LoadMTextIdFromXData(pl);
                
                if (!mtextId.HasValue || !mtextId.Value.IsValid)
                {
                    ed.WriteMessage($"\nKhông tìm thấy MText ObjectId trong XData của polyline Handle: {pl.ObjectId.Handle.Value}");
                    ed.WriteMessage($"\nVui lòng chạy lại lệnh TAN25_TANDUCTGOOD để tạo MText mới!");
                    return;
                }

                // Lấy điểm cuối của polyline
                Point3d endPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                
                // Tính vector pháp tuyến tại điểm cuối
                Vector3d tangent = GetTangentAtEnd(pl);
                Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0).GetNormal();
                
                // Tạo text chỉ hiển thị width (như trước)
                var widthText = $"{newWidth:F0}\"∅";
                
                // Vị trí text bên ngoài polyline
                var textPosition = endPoint + normal * (3.0 + newWidth);
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Mở MText để cập nhật
                        MText mtext = tr.GetObject(mtextId.Value, OpenMode.ForWrite) as MText;
                        if (mtext != null)
                        {
                            mtext.Contents = widthText;
                            mtext.Location = textPosition;
                            ed.WriteMessage($"\nĐã cập nhật MText trực tiếp với width mới: {newWidth:F1}");
                        }
                        else
                        {
                            ed.WriteMessage($"\nKhông thể mở MText với ObjectId: {mtextId.Value.Handle.Value}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nLỗi cập nhật MText: {ex.Message}");
                        ed.WriteMessage($"\nMText có thể đã bị xóa, sẽ tạo mới...");
                        
                        // Tạo MText mới nếu không tìm thấy
                        CreateMTextForPolyline(pl);
                        return;
                    }

                    tr.Commit();
                }
                
                // Refresh màn hình để hiển thị MText
                ed.Regen();
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi cập nhật MText: {ex.Message}");
            }
        }


        /// <summary>
        /// Tạo MText cho polyline sau khi kết thúc lệnh vẽ
        /// </summary>
        /// <param name="pl">Polyline đã vẽ</param>
        private static void CreateMTextForPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                // Load width từ XData của polyline
                double width = LoadWidthFromXData(pl);
                if (width <= 0)
                {
                    ed.WriteMessage($"\nKhông tìm thấy width trong XData của polyline Handle: {pl.ObjectId.Handle.Value}");
                    return;
                }

                // Lấy điểm cuối của polyline
                Point3d endPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                
                // Tính vector pháp tuyến tại điểm cuối
                Vector3d tangent = GetTangentAtEnd(pl);
                Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0).GetNormal();
                
                // Tạo text chỉ hiển thị width (như trước)
                var widthText = $"{width:F0}\"∅";
                
                // Vị trí text bên ngoài polyline
                var textPosition = endPoint + normal * (3.0 + width); // Cách polyline 5 inch
                
                // Tạo layer cho MText nếu chưa có
                ObjectId layerId = CreateOrGetLayer("M-ANNO-TAG-DUCT", 50); // Màu 50 theo chuẩn
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Tạo MText
                    MText mtext = new MText();
                    mtext.SetDatabaseDefaults();
                    mtext.Contents = widthText;
                    mtext.Location = textPosition;
                    mtext.Height = 4.5; // Chiều cao text
                    mtext.TextHeight = 4.5; // Đảm bảo height không bị override
                    mtext.ColorIndex = 50; // Màu 50 theo chuẩn
                    mtext.LayerId = layerId;
                    
                    // Đảm bảo height không bị override bởi layer
                    mtext.Color = Color.FromColorIndex(ColorMethod.ByAci, 50);
                    
                    // Thêm MText vào database
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);
                    
                    tr.Commit();
                    
                    // Lưu MText ObjectId vào XData của polyline
                    SaveTickInfoToDatabase(pl, width, mtext.ObjectId);
                }
                
                // Refresh màn hình để hiển thị MText
                ed.Regen();
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi tạo MText: {ex.Message}");
            }
        }

        /// <summary>
        /// Lấy vector tiếp tuyến tại điểm cuối của polyline
        /// </summary>
        /// <param name="pl">Polyline</param>
        /// <returns>Vector tiếp tuyến</returns>
        private static Vector3d GetTangentAtEnd(Autodesk.AutoCAD.DatabaseServices.Polyline pl)
        {
            try
            {
                if (pl.NumberOfVertices < 2) return new Vector3d(1, 0, 0);
                
                Point3d lastPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                Point3d secondLastPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 2);
                
                return (lastPoint - secondLastPoint).GetNormal();
            }
            catch
            {
                return new Vector3d(1, 0, 0);
            }
        }


        /// <summary>
        /// Lệnh edit width của polyline và tự động load lại tick + MText
        /// </summary>
        [CommandMethod("TAN25_HVAC_DUCT_SUPPLY_AIR_EDIT_WIDTH")]
        public static void EditWidthAndReload()
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                ed.WriteMessage("\n=== BẮT ĐẦU LỆNH TAN25_EDITWIDTH ===");

                // Chọn polyline
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline cần edit width: ");
                peo.SetRejectMessage("\nVui lòng chọn polyline!");
                peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);
                peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline2d), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl == null)
                    {
                        ed.WriteMessage("\nKhông thể load polyline!");
                        return;
                    }

                    // Load width hiện tại từ XData
                    double currentWidth = LoadWidthFromXData(pl);
                    if (currentWidth <= 0)
                    {
                        ed.WriteMessage("\nPolyline này không có width trong XData!");
                        return;
                    }

                    ed.WriteMessage($"\nWidth hiện tại: {currentWidth:F1}");

                    // Nhập width mới
                    PromptDoubleOptions pdo = new PromptDoubleOptions($"\nNhập width mới (hiện tại: {currentWidth:F1}): ");
                    pdo.AllowNegative = false;
                    pdo.AllowZero = false;
                    pdo.DefaultValue = currentWidth;

                    PromptDoubleResult pdr = ed.GetDouble(pdo);
                    if (pdr.Status != PromptStatus.OK) return;

                    double newWidth = pdr.Value;

                    // Cập nhật XData với width mới
                    pl.UpgradeOpen();
                    UpdateWidthInXData(pl, newWidth);

                    // Cập nhật width trong FilletOverrule
                    FilletOverrule.UpdateBlueWidth(newWidth);
                    FilletOverrule.AddAllowedPolylineWithWidth(pl.ObjectId, newWidth);

                    tr.Commit();

                    ed.WriteMessage($"\nĐã cập nhật width từ {currentWidth:F1} thành {newWidth:F1}");

                    // Cập nhật MText trực tiếp bằng ObjectId từ XData
                    UpdateMTextForPolyline(pl, newWidth);

                    // Load lại tick
                    LoadTempTicks();

                    ed.WriteMessage($"\nĐã load lại tick và MText với width mới: {newWidth:F1}");
                }

                // Refresh màn hình để hiển thị thay đổi
                ed.Regen();
                ed.WriteMessage("\n=== KẾT THÚC LỆNH TAN25_EDITWIDTH ===");
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi trong lệnh TAN25_EDITWIDTH: {ex.Message}");
            }
        }







        /// <summary>
        /// Lệnh tự động load tick sau khi break polyline
        /// </summary>
        [CommandMethod("TAN25_HVAC_DUC_AUTOBREAK")]
        public static void AutoBreakWithTick()
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                ed.WriteMessage("\n=== BẮT ĐẦU LỆNH TAN25_AUTOBREAK ===");
                ed.WriteMessage("\nHướng dẫn:");
                ed.WriteMessage("\n1. Chọn polyline cần break");
                ed.WriteMessage("\n2. Chọn điểm break thứ nhất");
                ed.WriteMessage("\n3. Chọn điểm break thứ hai");
                ed.WriteMessage("\n4. Hệ thống sẽ tự động load tick cho 2 polyline mới");

                // Chọn polyline cần break
                PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline cần break: ");
                peo.SetRejectMessage("\nVui lòng chọn polyline!");
                peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                double width = 0;
                ResultBuffer originalXData = null;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var originalPl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (originalPl == null)
                    {
                        ed.WriteMessage("\nKhông thể load polyline!");
                        return;
                    }

                    // Kiểm tra XData
                    width = LoadWidthFromXData(originalPl);
                    if (width <= 0)
                    {
                        ed.WriteMessage("\nPolyline này không có XData HVAC_DUCT_SUPPLY_AIR!");
                        return;
                    }

                    ed.WriteMessage($"\nWidth từ polyline gốc: {width:F1}");

                    // Lưu thông tin XData để sử dụng sau
                    originalXData = originalPl.XData;

                    // Xóa tick ở điểm đầu và cuối trước khi break
                    ed.WriteMessage("\nĐang xóa tick ở điểm đầu và cuối...");
                    
                    // Xóa MText của polyline gốc
                    DeleteMTextForPolyline(originalPl);
                    
                    // Xóa tick khỏi danh sách hiển thị
                    FilletOverrule.RemoveAllowedPolyline(per.ObjectId);

                    tr.Commit();
                }

                // Chọn điểm break thứ nhất
                PromptPointOptions ppo1 = new PromptPointOptions("\nChọn điểm break thứ nhất: ");
                PromptPointResult ppr1 = ed.GetPoint(ppo1);
                if (ppr1.Status != PromptStatus.OK) return;

                Point3d breakPoint1 = ppr1.Value;

                // Chọn điểm break thứ hai
                PromptPointOptions ppo2 = new PromptPointOptions("\nChọn điểm break thứ hai: ");
                PromptPointResult ppr2 = ed.GetPoint(ppo2);
                if (ppr2.Status != PromptStatus.OK) return;

                Point3d breakPoint2 = ppr2.Value;

                // Thực hiện break polyline
                ed.WriteMessage("\nĐang thực hiện break polyline...");
                
                // Sử dụng lệnh BREAK của AutoCAD
                ed.Command("BREAK", per.ObjectId, breakPoint1, breakPoint2);

                // Đợi một chút để AutoCAD hoàn thành lệnh break
                System.Threading.Thread.Sleep(500);

                // Tìm và xử lý 2 polyline mới
                ProcessNewPolylinesAfterBreak(per.ObjectId, width, originalXData);
                LoadTempTicks();

                ed.WriteMessage("\n=== KẾT THÚC LỆNH TAN25_AUTOBREAK ===");
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi auto break: {ex.Message}");
            }
        }

        /// <summary>
        /// Xử lý 2 polyline mới sau khi break với width giống block
        /// </summary>
        private static void ProcessNewPolylinesAfterBreakWithBlockWidth(ObjectId originalObjectId, double blockWidth, ResultBuffer originalXData, Editor ed)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                ed.WriteMessage("\nĐang tìm polyline mới sau break...");

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Tìm tất cả polyline trong ModelSpace
                    var allPolylines = new List<Autodesk.AutoCAD.DatabaseServices.Polyline>();
                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null)
                            {
                                allPolylines.Add(pl);
                            }
                        }
                    }

                    ed.WriteMessage($"\nTổng số polyline trong drawing: {allPolylines.Count}");

                    // Tìm polyline không có XData (có thể là polyline mới từ break)
                    var newPolylines = new List<Autodesk.AutoCAD.DatabaseServices.Polyline>();
                    
                    foreach (var pl in allPolylines)
                    {
                        double currentWidth = LoadWidthFromXData(pl);
                        if (currentWidth <= 0) // Không có XData
                        {
                            newPolylines.Add(pl);
                            ed.WriteMessage($"\nTìm thấy polyline không có XData: Handle {pl.ObjectId.Handle.Value}");
                        }
                    }

                    ed.WriteMessage($"\nTìm thấy {newPolylines.Count} polyline không có XData");

                    // Xử lý tất cả polyline không có XData với width giống block
                    int processedCount = 0;
                    foreach (var newPl in newPolylines)
                    {
                        try
                        {
                            // Tạo XData mới cho polyline này với width giống block
                            newPl.UpgradeOpen();
                            
                            // Tạo XData mới với width từ block
                            SaveTickInfoToDatabase(newPl, blockWidth);
                            
                            // Tạo MText mới cho polyline này
                            CreateMTextForPolyline(newPl);
                            
                            // Thêm vào danh sách hiển thị tick với width giống block
                            FilletOverrule.AddAllowedPolylineWithWidth(newPl.ObjectId, blockWidth);
                            
                            processedCount++;
                            ed.WriteMessage($"\nĐã tạo XData mới cho polyline Handle: {newPl.ObjectId.Handle.Value} với width: {blockWidth:F1}");
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nLỗi xử lý polyline Handle {newPl.ObjectId.Handle.Value}: {ex.Message}");
                        }
                    }

                    tr.Commit();
                    
                    if (processedCount > 0)
                    {
                        ed.WriteMessage($"\nĐã xử lý {processedCount} polyline mới với width giống block!");
                        
                        // Load lại tick
                        LoadTempTicks();
                        ed.WriteMessage("\nĐã load lại tick!");
                    }
                    else
                    {
                        ed.WriteMessage("\nKhông tìm thấy polyline mới để xử lý!");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed2 = doc.Editor;
                ed2.WriteMessage($"\nLỗi xử lý polyline mới: {ex.Message}");
            }
        }

        /// <summary>
        /// Xử lý 2 polyline mới sau khi break
        /// </summary>
        private static void ProcessNewPolylinesAfterBreak(ObjectId originalObjectId, double originalWidth, ResultBuffer originalXData)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                ed.WriteMessage("\nĐang tìm polyline mới sau break...");

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Tìm tất cả polyline trong ModelSpace
                    var allPolylines = new List<Autodesk.AutoCAD.DatabaseServices.Polyline>();
                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null)
                            {
                                allPolylines.Add(pl);
                            }
                        }
                    }

                    ed.WriteMessage($"\nTổng số polyline trong drawing: {allPolylines.Count}");

                    // Tìm polyline không có XData (có thể là polyline mới từ break)
                    var newPolylines = new List<Autodesk.AutoCAD.DatabaseServices.Polyline>();
                    
                    foreach (var pl in allPolylines)
                    {
                        double currentWidth = LoadWidthFromXData(pl);
                        if (currentWidth <= 0) // Không có XData
                        {
                            newPolylines.Add(pl);
                            ed.WriteMessage($"\nTìm thấy polyline không có XData: Handle {pl.ObjectId.Handle.Value}");
                        }
                    }

                    ed.WriteMessage($"\nTìm thấy {newPolylines.Count} polyline không có XData");

                    // Xử lý tất cả polyline không có XData
                    int processedCount = 0;
                    foreach (var newPl in newPolylines)
                    {
                        try
                        {
                            // Tạo XData mới cho polyline này (không copy từ polyline gốc)
                            newPl.UpgradeOpen();
                            
                            // Tạo XData mới với width từ polyline gốc
                            SaveTickInfoToDatabase(newPl, originalWidth);
                            
                            // Tạo MText mới cho polyline này
                            CreateMTextForPolyline(newPl);
                            
                            // Thêm vào danh sách hiển thị tick
                            FilletOverrule.AddAllowedPolylineWithWidth(newPl.ObjectId, originalWidth);
                            
                            processedCount++;
                            ed.WriteMessage($"\nĐã tạo XData mới cho polyline Handle: {newPl.ObjectId.Handle.Value}");
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nLỗi xử lý polyline Handle {newPl.ObjectId.Handle.Value}: {ex.Message}");
                        }
                    }

                    tr.Commit();
                    
                    if (processedCount > 0)
                    {
                        ed.WriteMessage($"\nĐã xử lý {processedCount} polyline mới!");
                        
                        // Load lại tick
                        LoadTempTicks();
                        ed.WriteMessage("\nĐã load lại tick!");
                    }
                    else
                    {
                        ed.WriteMessage("\nKhông tìm thấy polyline mới để xử lý!");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi xử lý polyline mới: {ex.Message}");
            }
        }


        /// <summary>
        /// Lệnh thêm block GE_Y_CONN vào polyline
        /// </summary>
        [CommandMethod("TAN25_HVAC_DUCT_ADD_Y_CONN")]
        public static void AddYConnBlock()
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                ed.WriteMessage("\n=== BẮT ĐẦU LỆNH TAN25_HVAC_DUCT_ADD_Y_CONN ===");
                ed.WriteMessage("\nHướng dẫn:");
                ed.WriteMessage("\n1. Click chọn điểm trên polyline để đặt Y-Connector");
                ed.WriteMessage("\n2. Hệ thống sẽ tự động tìm polyline và break tại vị trí đó");

                // Chọn điểm trên polyline (tự động tìm polyline)
                PromptPointOptions ppo = new PromptPointOptions("\nClick chọn điểm trên polyline để đặt Y-Connector: ");
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                Point3d userPoint = ppr.Value;
                
                // Tìm polyline gần nhất với điểm chọn
                ObjectId plId = FindNearestPolylineWithXData(userPoint, db, ed);
                if (plId.IsNull)
                {
                    ed.WriteMessage("\nKhông tìm thấy polyline có XData HVAC_DUCT gần điểm chọn!");
                    return;
                }
                
                // Kiểm tra XData của polyline và tìm điểm chính xác
                double width = 0;
                Point3d insertPoint = userPoint; // Default to user point
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(plId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl == null)
                    {
                        ed.WriteMessage("\nKhông thể load polyline!");
                        return;
                    }

                    // Tìm điểm chính xác trên polyline gần nhất với điểm user chọn
                    insertPoint = GetClosestPointOnPolyline(pl, userPoint);
                    ed.WriteMessage($"\nĐiểm user chọn: ({userPoint.X:F2}, {userPoint.Y:F2})");
                    ed.WriteMessage($"\nĐiểm chính xác trên polyline: ({insertPoint.X:F2}, {insertPoint.Y:F2})");

                    width = LoadWidthFromXData(pl);
                    if (width <= 0)
                    {
                        ed.WriteMessage("\nPolyline này không có XData HVAC_DUCT_SUPPLY_AIR!");
                        return;
                    }

                    ed.WriteMessage($"\nWidth từ polyline: {width:F1}");

                    // Load block GE_Y_CONN từ file
                    ed.WriteMessage("\nĐang load block GE_Y_CONN...");
                    ObjectId blockId = LoadYConnBlock(db, width, ed);
                    if (blockId.IsNull)
                    {
                        ed.WriteMessage("\nKhông thể load block GE_Y_CONN!");
                        return;
                    }
                    ed.WriteMessage($"\nĐã load block GE_Y_CONN thành công! Handle: {blockId.Handle.Value}");

                    // Chèn block vào điểm đã chọn
                    ed.WriteMessage("\nĐang chèn block vào ModelSpace...");
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    BlockReference blockRef = new BlockReference(insertPoint, blockId);
                    blockRef.SetDatabaseDefaults();
                    blockRef.Color = Color.FromColorIndex(ColorMethod.ByAci, 4); // Màu xanh sáng
                    
                    // Tính góc quay dựa trên hướng của polyline tại điểm chèn
                    double rotation = GetPolylineRotationAtPoint(pl, insertPoint);
                    blockRef.Rotation = rotation;
                    
                        // Tính 2 điểm kết nối của block (đầu và cuối)
                        Vector3d direction = new Vector3d(Math.Cos(rotation), Math.Sin(rotation), 0);
                        double blockLength = width * 0.6; // Chiều dài block dựa trên width (tăng lên)
                        Point3d blockStart = insertPoint - direction * blockLength;
                        Point3d blockEnd = insertPoint + direction * blockLength;
                        
                        // Tìm điểm kết nối chính xác trên polyline
                        Point3d connectionPoint1 = GetClosestPointOnPolyline(pl, blockStart);
                        Point3d connectionPoint2 = GetClosestPointOnPolyline(pl, blockEnd);
                        
                        // Đảm bảo 2 điểm kết nối không trùng nhau
                        if (connectionPoint1.DistanceTo(connectionPoint2) < 0.1)
                        {
                            // Nếu 2 điểm quá gần, tạo 2 điểm cách nhau
                            Vector3d perpDirection = new Vector3d(-direction.Y, direction.X, 0);
                            connectionPoint1 = insertPoint - perpDirection * (width * 0.3);
                            connectionPoint2 = insertPoint + perpDirection * (width * 0.3);
                        }
                    
                    ed.WriteMessage($"\nĐiểm kết nối 1: ({connectionPoint1.X:F2}, {connectionPoint1.Y:F2})");
                    ed.WriteMessage($"\nĐiểm kết nối 2: ({connectionPoint2.X:F2}, {connectionPoint2.Y:F2})");
                    
                    ed.WriteMessage($"\nBlock sẽ được chèn tại: ({insertPoint.X:F2}, {insertPoint.Y:F2}) với góc {rotation * 180 / Math.PI:F1}°");
                    ed.WriteMessage($"\nĐiểm đầu block: ({blockStart.X:F2}, {blockStart.Y:F2})");
                    ed.WriteMessage($"\nĐiểm cuối block: ({blockEnd.X:F2}, {blockEnd.Y:F2})");

                    btr.AppendEntity(blockRef);
                    tr.AddNewlyCreatedDBObject(blockRef, true);
                    
                    ed.WriteMessage($"\nBlock đã được thêm vào ModelSpace! Handle: {blockRef.ObjectId.Handle.Value}");

                    // Break polyline tại vị trí chèn block
                    ed.WriteMessage("\nĐang break polyline tại vị trí block...");
                    
                    // Lưu ObjectId trước khi break
                    ObjectId originalPlId = pl.ObjectId;
                    
                    tr.Commit(); // Commit transaction trước khi break
                    
                    // Break polyline tại 2 điểm kết nối của block
                    BreakPolylineAtBlockPoints(originalPlId, connectionPoint1, connectionPoint2, width, ed);

                    ed.WriteMessage($"\nĐã thêm Y-Connector và break polyline thành công tại điểm ({insertPoint.X:F2}, {insertPoint.Y:F2}) với góc {rotation * 180 / Math.PI:F1}°");
                }

                ed.WriteMessage("\n=== KẾT THÚC LỆNH TAN25_HVAC_DUCT_ADD_Y_CONN ===");
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi thêm Y-Connector: {ex.Message}");
            }
        }

        /// <summary>
        /// Load block GE_Y_CONN từ file
        /// </summary>
        private static ObjectId LoadYConnBlock(Database db, double width, Editor ed)
        {
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    
                    // Kiểm tra block đã tồn tại chưa
                    if (bt.Has("GE_Y_CONN"))
                    {
                        return bt["GE_Y_CONN"];
                    }

                    // Đường dẫn file block (có thể thay đổi theo nhu cầu)
                    string blockFilePath = GetBlockFilePath();
                    ed.WriteMessage($"\nĐang tìm file block tại: {blockFilePath}");
                    
                    if (string.IsNullOrEmpty(blockFilePath) || !System.IO.File.Exists(blockFilePath))
                    {
                        ed.WriteMessage("\nKhông tìm thấy file block, sẽ tạo block mặc định...");
                        // Nếu không tìm thấy file block, tạo block mặc định
                        return CreateDefaultYConnBlock(db, width, tr, ed);
                    }
                    
                    ed.WriteMessage($"\nTìm thấy file block: {blockFilePath}");

                    // Load block từ file
                    Database sourceDb = new Database(false, true);
                    try
                    {
                        sourceDb.ReadDwgFile(blockFilePath, FileShare.Read, true, "");
                        
                        using (Transaction sourceTr = sourceDb.TransactionManager.StartTransaction())
                        {
                            BlockTable sourceBt = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                            
                            if (sourceBt.Has("GE_Y_CONN"))
                            {
                                // Copy block từ file nguồn
                                ObjectId sourceBlockId = sourceBt["GE_Y_CONN"];
                                BlockTableRecord sourceBtr = sourceTr.GetObject(sourceBlockId, OpenMode.ForRead) as BlockTableRecord;
                                
                                // Tạo block mới trong database hiện tại
                                BlockTableRecord newBtr = new BlockTableRecord();
                                newBtr.Name = "GE_Y_CONN";
                                newBtr.Origin = sourceBtr.Origin;
                                
                                // Copy các entity từ block gốc
                                foreach (ObjectId entityId in sourceBtr)
                                {
                                    Entity entity = sourceTr.GetObject(entityId, OpenMode.ForRead) as Entity;
                                    if (entity != null)
                                    {
                                        Entity clonedEntity = entity.Clone() as Entity;
                                        if (clonedEntity != null)
                                        {
                                            clonedEntity.SetDatabaseDefaults();
                                            newBtr.AppendEntity(clonedEntity);
                                        }
                                    }
                                }
                                
                                // Thêm block vào BlockTable
                                bt.UpgradeOpen();
                                ObjectId blockId = bt.Add(newBtr);
                                tr.AddNewlyCreatedDBObject(newBtr, true);
                                
                                sourceTr.Commit();
                                tr.Commit();
                                
                                return blockId;
                            }
                            else
                            {
                                sourceTr.Commit();
                                // Nếu không tìm thấy block trong file, tạo block mặc định
                                return CreateDefaultYConnBlock(db, width, tr, ed);
                            }
                        }
                    }
                    finally
                    {
                        sourceDb.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor editor = doc.Editor;
                editor.WriteMessage($"\nLỗi load block GE_Y_CONN từ file: {ex.Message}");
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// Lấy đường dẫn file block
        /// </summary>
        private static string GetBlockFilePath()
        {
            try
            {
                // Đường dẫn file block cụ thể
                string[] possiblePaths = {
                    @"D:\Blocks\Mechanical\GE_Y_CONN.dwg",  // Đường dẫn chính
                    @"C:\HVAC_Blocks\GE_Y_CONN.dwg",
                    @"D:\HVAC_Blocks\GE_Y_CONN.dwg",
                    @"C:\Program Files\HVAC_DUCT\Blocks\GE_Y_CONN.dwg",
                    System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Blocks", "GE_Y_CONN.dwg"),
                    System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "GE_Y_CONN.dwg")
                };

                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor editor = doc.Editor;
                editor.WriteMessage("\nĐang tìm file block trong các đường dẫn:");
                
                foreach (string path in possiblePaths)
                {
                    editor.WriteMessage($"\n- Kiểm tra: {path}");
                    if (System.IO.File.Exists(path))
                    {
                        editor.WriteMessage($"\n✓ Tìm thấy file: {path}");
                        return path;
                    }
                    else
                    {
                        editor.WriteMessage($"\n✗ Không tìm thấy: {path}");
                    }
                }

                editor.WriteMessage("\nKhông tìm thấy file block nào!");
                return string.Empty;
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor editor = doc.Editor;
                editor.WriteMessage($"\nLỗi khi tìm file block: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Tạo block GE_Y_CONN mặc định (fallback)
        /// </summary>
        private static ObjectId CreateDefaultYConnBlock(Database db, double width, Transaction tr, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nĐang tạo block GE_Y_CONN mặc định...");
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                
                // Tạo block mới
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = "GE_Y_CONN";
                btr.Origin = Point3d.Origin;

                // Tạo layer cho block
                ObjectId layerId = CreateOrGetLayer("M-DUCT-FITTING", 4);

                // Tạo hình dạng Y-Connector
                double scale = width / 12.0; // Scale dựa trên width
                double size = 6.0 * scale; // Kích thước cơ bản

                // Vẽ 3 đường thẳng tạo hình Y
                // Đường chính (dọc)
                Line mainLine = new Line(
                    new Point3d(0, -size, 0),
                    new Point3d(0, size, 0)
                );
                mainLine.SetDatabaseDefaults();
                mainLine.Color = Color.FromColorIndex(ColorMethod.ByAci, 4);
                mainLine.LayerId = layerId;
                btr.AppendEntity(mainLine);

                // Nhánh trái
                Line leftBranch = new Line(
                    new Point3d(0, 0, 0),
                    new Point3d(-size * 0.7, size * 0.7, 0)
                );
                leftBranch.SetDatabaseDefaults();
                leftBranch.Color = Color.FromColorIndex(ColorMethod.ByAci, 4);
                leftBranch.LayerId = layerId;
                btr.AppendEntity(leftBranch);

                // Nhánh phải
                Line rightBranch = new Line(
                    new Point3d(0, 0, 0),
                    new Point3d(size * 0.7, size * 0.7, 0)
                );
                rightBranch.SetDatabaseDefaults();
                rightBranch.Color = Color.FromColorIndex(ColorMethod.ByAci, 4);
                rightBranch.LayerId = layerId;
                btr.AppendEntity(rightBranch);

                // Thêm block vào BlockTable
                bt.UpgradeOpen();
                ObjectId blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                
                ed.WriteMessage($"\nĐã tạo block GE_Y_CONN mặc định thành công! Handle: {blockId.Handle.Value}");

                return blockId;
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor editor = doc.Editor;
                editor.WriteMessage($"\nLỗi tạo block GE_Y_CONN mặc định: {ex.Message}");
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// Break polyline tại 2 điểm kết nối của block
        /// </summary>
        private static void BreakPolylineAtBlockPoints(ObjectId plId, Point3d connectionPoint1, Point3d connectionPoint2, double width, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nBắt đầu break polyline tại 2 điểm kết nối của block...");
                
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                
                ed.WriteMessage($"\nĐiểm kết nối 1: ({connectionPoint1.X:F2}, {connectionPoint1.Y:F2})");
                ed.WriteMessage($"\nĐiểm kết nối 2: ({connectionPoint2.X:F2}, {connectionPoint2.Y:F2})");
                
                // Lưu XData của polyline gốc trước khi break
                ResultBuffer originalXData = null;
                Point3d breakPoint1 = connectionPoint1;
                Point3d breakPoint2 = connectionPoint2;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(plId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl != null)
                    {
                        originalXData = pl.XData;
                        
                        ed.WriteMessage($"\nĐiểm break 1: ({breakPoint1.X:F2}, {breakPoint1.Y:F2})");
                        ed.WriteMessage($"\nĐiểm break 2: ({breakPoint2.X:F2}, {breakPoint2.Y:F2})");
                        
                        // Xóa MText của polyline gốc
                        DeleteMTextForPolyline(pl);
                        
                        // Xóa tick khỏi danh sách hiển thị
                        FilletOverrule.RemoveAllowedPolyline(plId);
                    }
                    tr.Commit();
                }

                // Thực hiện break polyline tại 2 điểm
                ed.WriteMessage("\nĐang thực hiện break polyline tại 2 điểm...");
                
                try
                {
                    // Break polyline tại điểm đầu block (không kéo dài)
                    ed.WriteMessage("\nBreak tại điểm đầu block...");
                    ed.Command("BREAK", plId, breakPoint1, breakPoint1);
                    System.Threading.Thread.Sleep(500);
                    
                    // Tìm polyline mới sau break đầu tiên
                    ObjectId newPlId = FindNewPolylineAfterBreak(plId, db);
                    if (!newPlId.IsNull)
                    {
                        // Break polyline thứ 2 tại điểm cuối block
                        ed.WriteMessage("\nBreak tại điểm cuối block...");
                        ed.Command("BREAK", newPlId, breakPoint2, breakPoint2);
                        System.Threading.Thread.Sleep(500);
                    }
                    
                    // Kéo dài polyline đến block sau khi break
                    ed.WriteMessage("\nKéo dài polyline đến block...");
                    ExtendPolylinesToBlock(breakPoint1, breakPoint2, ed);
                    
                    // Điều chỉnh polyline để kết nối chính xác với block
                    AdjustPolylinesToBlock(breakPoint1, breakPoint2, ed);
                    
                    ed.WriteMessage("\nLệnh BREAK đã được thực hiện tại 2 điểm");

                    // Đợi một chút để AutoCAD hoàn thành lệnh break
                    System.Threading.Thread.Sleep(1000);

                    // Tìm và xử lý 2 polyline mới với width giống block
                    ed.WriteMessage("\nĐang tìm polyline mới sau break...");
                    ProcessNewPolylinesAfterBreakWithBlockWidth(plId, width, originalXData, ed);
                    
                    ed.WriteMessage("\nĐã break polyline thành công tại 2 điểm!");
                }
                catch (System.Exception breakEx)
                {
                    ed.WriteMessage($"\nLỗi khi thực hiện lệnh BREAK: {breakEx.Message}");
                    ed.WriteMessage("\nSẽ thử cách khác...");
                    
                    // Fallback: Không break polyline, chỉ thêm block
                    ed.WriteMessage("\nChỉ thêm block mà không break polyline");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi break polyline: {ex.Message}");
            }
        }

        /// <summary>
        /// Kéo dài polyline đến block sau khi break
        /// </summary>
        private static void ExtendPolylinesToBlock(Point3d point1, Point3d point2, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nĐang kéo dài polyline đến block...");
                
                // Tìm tất cả polyline mới (không có XData)
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                var newPolylineIds = new List<ObjectId>();
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null && LoadWidthFromXData(pl) <= 0) // Không có XData
                            {
                                newPolylineIds.Add(objId);
                            }
                        }
                    }
                    tr.Commit();
                }
                
                ed.WriteMessage($"\nTìm thấy {newPolylineIds.Count} polyline mới để kéo dài");
                
                // Kéo dài từng polyline đến điểm gần nhất
                foreach (var plId in newPolylineIds)
                {
                    try
                    {
                        // Tìm điểm gần nhất trên polyline với mỗi điểm block
                        Point3d closestToPoint1 = GetClosestPointOnPolylineById(plId, point1, db);
                        Point3d closestToPoint2 = GetClosestPointOnPolylineById(plId, point2, db);
                        
                        // Kéo dài đến điểm gần nhất
                        ed.Command("EXTEND", plId, "", closestToPoint1, closestToPoint2);
                        ed.WriteMessage($"\nĐã kéo dài polyline Handle: {plId.Handle.Value}");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nLỗi kéo dài polyline Handle {plId.Handle.Value}: {ex.Message}");
                    }
                }
                
                ed.WriteMessage("\nĐã kéo dài tất cả polyline đến block");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi kéo dài polyline: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm điểm gần nhất trên polyline theo ObjectId
        /// </summary>
        private static Point3d GetClosestPointOnPolylineById(ObjectId plId, Point3d point, Database db)
        {
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(plId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl != null)
                    {
                        return GetClosestPointOnPolyline(pl, point);
                    }
                    tr.Commit();
                }
            }
            catch
            {
                // Fallback to original point
            }
            return point;
        }

        /// <summary>
        /// Điều chỉnh polyline để kết nối chính xác với block
        /// </summary>
        private static void AdjustPolylinesToBlock(Point3d point1, Point3d point2, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nĐang điều chỉnh polyline để kết nối chính xác với block...");
                
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                var newPolylineIds = new List<ObjectId>();
                
                // Tìm tất cả polyline mới (không có XData)
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null && LoadWidthFromXData(pl) <= 0) // Không có XData
                            {
                                newPolylineIds.Add(objId);
                            }
                        }
                    }
                    tr.Commit();
                }
                
                ed.WriteMessage($"\nTìm thấy {newPolylineIds.Count} polyline mới để điều chỉnh");
                
                // Điều chỉnh từng polyline
                foreach (var plId in newPolylineIds)
                {
                    try
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            var pl = tr.GetObject(plId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null)
                            {
                                // Tìm điểm gần nhất với mỗi điểm block
                                Point3d closestToPoint1 = GetClosestPointOnPolyline(pl, point1);
                                Point3d closestToPoint2 = GetClosestPointOnPolyline(pl, point2);
                                
                                // Tính khoảng cách từ polyline đến block
                                double distToPoint1 = closestToPoint1.DistanceTo(point1);
                                double distToPoint2 = closestToPoint2.DistanceTo(point2);
                                
                                // Nếu polyline gần với điểm 1, kéo dài đến điểm 1
                                if (distToPoint1 < distToPoint2)
                                {
                                    // Kéo dài polyline đến điểm 1
                                    ExtendPolylineToPoint(pl, point1);
                                    ed.WriteMessage($"\nĐã kéo dài polyline Handle: {plId.Handle.Value} đến điểm 1");
                                }
                                else
                                {
                                    // Kéo dài polyline đến điểm 2
                                    ExtendPolylineToPoint(pl, point2);
                                    ed.WriteMessage($"\nĐã kéo dài polyline Handle: {plId.Handle.Value} đến điểm 2");
                                }
                            }
                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nLỗi điều chỉnh polyline Handle {plId.Handle.Value}: {ex.Message}");
                    }
                }
                
                ed.WriteMessage("\nĐã điều chỉnh tất cả polyline để kết nối với block");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi điều chỉnh polyline: {ex.Message}");
            }
        }

        /// <summary>
        /// Kéo dài polyline đến điểm cụ thể
        /// </summary>
        private static void ExtendPolylineToPoint(Autodesk.AutoCAD.DatabaseServices.Polyline pl, Point3d targetPoint)
        {
            try
            {
                // Tìm điểm cuối gần nhất với target point
                Point3d startPoint = pl.GetPoint3dAt(0);
                Point3d endPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                
                Point3d closestEnd = startPoint.DistanceTo(targetPoint) < endPoint.DistanceTo(targetPoint) ? startPoint : endPoint;
                
                // Nếu điểm cuối gần nhất là điểm đầu, thêm điểm mới vào đầu
                if (closestEnd == startPoint)
                {
                    pl.AddVertexAt(0, new Point2d(targetPoint.X, targetPoint.Y), 0, 0, 0);
                }
                else
                {
                    // Thêm điểm mới vào cuối
                    pl.AddVertexAt(pl.NumberOfVertices, new Point2d(targetPoint.X, targetPoint.Y), 0, 0, 0);
                }
            }
            catch (System.Exception ex)
            {
                // Log error but don't throw
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage($"\nLỗi kéo dài polyline: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm polyline mới sau break
        /// </summary>
        private static ObjectId FindNewPolylineAfterBreak(ObjectId originalPlId, Database db)
        {
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Tìm polyline mới (không có XData)
                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            if (objId != originalPlId) // Không phải polyline gốc
                            {
                                var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                                if (pl != null && LoadWidthFromXData(pl) <= 0) // Không có XData
                                {
                                    tr.Commit();
                                    return objId;
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            catch
            {
                // Ignore errors
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// Break polyline tại điểm chỉ định
        /// </summary>
        private static void BreakPolylineAtPoint(ObjectId plId, Point3d breakPoint, double width, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nBắt đầu break polyline...");
                
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                
                // Lưu XData của polyline gốc trước khi break
                ResultBuffer originalXData = null;
                Point3d breakPointExact = breakPoint;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(plId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                    if (pl != null)
                    {
                        originalXData = pl.XData;
                        
                        // Tìm điểm chính xác trên polyline
                        breakPointExact = GetClosestPointOnPolyline(pl, breakPoint);
                        ed.WriteMessage($"\nĐiểm break chính xác trên polyline: ({breakPointExact.X:F2}, {breakPointExact.Y:F2})");
                        
                        // Xóa MText của polyline gốc
                        DeleteMTextForPolyline(pl);
                        
                        // Xóa tick khỏi danh sách hiển thị
                        FilletOverrule.RemoveAllowedPolyline(plId);
                    }
                    tr.Commit();
                }

                // Thực hiện break polyline
                ed.WriteMessage("\nĐang thực hiện break polyline...");
                
                try
                {
                    // Sử dụng lệnh BREAK của AutoCAD với điểm chính xác
                    ed.Command("BREAK", plId, breakPointExact, breakPointExact);
                    ed.WriteMessage("\nLệnh BREAK đã được thực hiện");

                    // Đợi một chút để AutoCAD hoàn thành lệnh break
                    System.Threading.Thread.Sleep(1000);

                    // Tìm và xử lý 2 polyline mới
                    ed.WriteMessage("\nĐang tìm polyline mới sau break...");
                    ProcessNewPolylinesAfterBreak(plId, width, originalXData);
                    
                    ed.WriteMessage("\nĐã break polyline thành công!");
                }
                catch (System.Exception breakEx)
                {
                    ed.WriteMessage($"\nLỗi khi thực hiện lệnh BREAK: {breakEx.Message}");
                    ed.WriteMessage("\nSẽ thử cách khác...");
                    
                    // Fallback: Không break polyline, chỉ thêm block
                    ed.WriteMessage("\nChỉ thêm block mà không break polyline");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi break polyline: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm polyline gần nhất có XData HVAC_DUCT
        /// </summary>
        private static ObjectId FindNearestPolylineWithXData(Point3d point, Database db, Editor ed)
        {
            try
            {
                ed.WriteMessage("\nĐang tìm polyline gần nhất có XData HVAC_DUCT...");
                
                ObjectId nearestPlId = ObjectId.Null;
                double minDistance = double.MaxValue;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                        {
                            var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                            if (pl != null)
                            {
                                // Kiểm tra có XData HVAC_DUCT không
                                double width = LoadWidthFromXData(pl);
                                if (width > 0)
                                {
                                    // Tính khoảng cách từ điểm đến polyline
                                    Point3d closestPoint = GetClosestPointOnPolyline(pl, point);
                                    double distance = point.DistanceTo(closestPoint);
                                    
                                    if (distance < minDistance)
                                    {
                                        minDistance = distance;
                                        nearestPlId = objId;
                                        ed.WriteMessage($"\nTìm thấy polyline gần hơn: Handle {objId.Handle.Value}, khoảng cách: {distance:F2}");
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                
                if (!nearestPlId.IsNull)
                {
                    ed.WriteMessage($"\nĐã tìm thấy polyline gần nhất: Handle {nearestPlId.Handle.Value}, khoảng cách: {minDistance:F2}");
                }
                else
                {
                    ed.WriteMessage("\nKhông tìm thấy polyline nào có XData HVAC_DUCT!");
                }
                
                return nearestPlId;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi tìm polyline: {ex.Message}");
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// Tìm điểm gần nhất trên polyline
        /// </summary>
        private static Point3d GetClosestPointOnPolyline(Autodesk.AutoCAD.DatabaseServices.Polyline pl, Point3d point)
        {
            try
            {
                double minDist = double.MaxValue;
                Point3d closestPoint = point;

                for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                {
                    Point3d start = pl.GetPoint3dAt(i);
                    Point3d end = pl.GetPoint3dAt(i + 1);
                    
                    // Tính điểm gần nhất trên segment
                    Point3d closestOnSegment = GetClosestPointOnLineSegment(point, start, end);
                    double dist = point.DistanceTo(closestOnSegment);
                    
                    if (dist < minDist)
                    {
                        minDist = dist;
                        closestPoint = closestOnSegment;
                    }
                }

                return closestPoint;
            }
            catch
            {
                return point; // Fallback to original point
            }
        }

        /// <summary>
        /// Tính điểm gần nhất trên đoạn thẳng
        /// </summary>
        private static Point3d GetClosestPointOnLineSegment(Point3d point, Point3d lineStart, Point3d lineEnd)
        {
            Vector3d line = lineEnd - lineStart;
            Vector3d pointToStart = point - lineStart;
            
            double lineLength = line.Length;
            if (lineLength == 0) return lineStart;
            
            double t = Math.Max(0, Math.Min(1, pointToStart.DotProduct(line) / (lineLength * lineLength)));
            return lineStart + t * line;
        }

        /// <summary>
        /// Tính góc quay của polyline tại một điểm
        /// </summary>
        private static double GetPolylineRotationAtPoint(Autodesk.AutoCAD.DatabaseServices.Polyline pl, Point3d point)
        {
            try
            {
                // Tìm segment gần nhất với điểm
                double minDist = double.MaxValue;
                int nearestSegment = 0;

                for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                {
                    Point3d start = pl.GetPoint3dAt(i);
                    Point3d end = pl.GetPoint3dAt(i + 1);
                    
                    // Tính khoảng cách từ điểm đến segment
                    double dist = GetDistanceToLineSegment(point, start, end);
                    if (dist < minDist)
                    {
                        minDist = dist;
                        nearestSegment = i;
                    }
                }

                // Tính vector hướng của segment
                Point3d startPt = pl.GetPoint3dAt(nearestSegment);
                Point3d endPt = pl.GetPoint3dAt(nearestSegment + 1);
                Vector3d direction = endPt - startPt;
                
                // Tính góc quay
                return Math.Atan2(direction.Y, direction.X);
            }
            catch
            {
                return 0.0; // Góc mặc định
            }
        }

        /// <summary>
        /// Tính khoảng cách từ điểm đến đoạn thẳng
        /// </summary>
        private static double GetDistanceToLineSegment(Point3d point, Point3d lineStart, Point3d lineEnd)
        {
            Vector3d line = lineEnd - lineStart;
            Vector3d pointToStart = point - lineStart;
            
            double lineLength = line.Length;
            if (lineLength == 0) return pointToStart.Length;
            
            double t = Math.Max(0, Math.Min(1, pointToStart.DotProduct(line) / (lineLength * lineLength)));
            Point3d projection = lineStart + t * line;
            
            return point.DistanceTo(projection);
        }

        /// <summary>
        /// Lệnh load tick từ database
        /// </summary>
        [CommandMethod("TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP")]
        public static void LoadTempTicks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\n=== BẮT ĐẦU LỆNH TAN25_LOADTEMP ===");

            try
            {
                // Đảm bảo overrule được kích hoạt
                if (_overrule == null)
                {
                    _overrule = new FilletOverrule();
                    Overrule.AddOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule, true);
                    Overrule.Overruling = true;
                    ed.WriteMessage("\nĐã kích hoạt tick overrule!");
                }

                ed.WriteMessage("\nĐang tìm polylines có tick trong database...");

                // Đăng ký ứng dụng trước khi đọc XData
                RegisterApplication(db);

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    int totalPolylines = 0;
                    int count = 0;
                    foreach (ObjectId objId in btr)
                    {
                        try
                        {
                            if (objId.ObjectClass == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)))
                            {
                                totalPolylines++;
                                var pl = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                                
                                // Kiểm tra XData có chứa thông tin tick không
                                var blueWidth = LoadWidthFromXData(pl);
                                if (blueWidth > 0)
                                {
                                    // Thêm vào danh sách hiển thị với width riêng
                                    FilletOverrule.AddAllowedPolylineWithWidth(objId, blueWidth);
                                    count++;
                                    ed.WriteMessage($"\nĐã load tick cho Handle: {objId.Handle.Value}, BlueWidth: {blueWidth:F1}");
                                }
                                else
                                {
                                    ed.WriteMessage($"\nPolyline Handle: {objId.Handle.Value} không có XData HVAC_DUCT");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nLỗi khi xử lý polyline: {ex.Message}");
                            continue;
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nTổng số polyline trong drawing: {totalPolylines}");
                    ed.WriteMessage($"\nĐã tải lại tick cho {count} polyline từ database!");
                }

                // Refresh màn hình để hiển thị tick
                doc.Editor.Regen();
                ed.WriteMessage("\n=== KẾT THÚC LỆNH TAN25_LOADTEMP ===");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nLỗi: {ex.Message}");
                ed.WriteMessage("\n=== LỖI TRONG LỆNH TAN25_LOADTEMP ===");
            }
        }

        /// <summary>
        /// DrawJig để vẽ polyline tương tác cho duct
        /// </summary>
        public class DuctDrawJig : Autodesk.AutoCAD.EditorInput.DrawJig
        {
            #region Constants
            private const double TICK_SPACING = 4.0;
            private const double RED_LENGTH = 2.0;
            #endregion

            #region Private Fields
            private Autodesk.AutoCAD.DatabaseServices.Polyline _polyline;
            private double _blueWidth;
            private Point3d _currentPoint;
            private bool _hasPreviewPoint = false;
            #endregion

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
                // ẨN polyline chính - chỉ vẽ polyline fit màu vàng
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
                    
                    // Thêm điểm preview nếu có
                    if (_hasPreviewPoint)
                    {
                        points.Add(_currentPoint);
                    }

                    if (points.Count < 2) return;

                    // Tạo polyline fit với bo góc giống polyline tạm
                    FilletOverrule overrule = new FilletOverrule();
                    var tempPl = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    
                    // Copy điểm vào polyline tạm
                    for (int i = 0; i < points.Count; i++)
                    {
                        tempPl.AddVertexAt(i, new Point2d(points[i].X, points[i].Y), 0, 0, 0);
                    }
                    
                    List<Point3d> fitPoints = overrule.CreateTempPolylinePoints(tempPl);
                    tempPl.Dispose();

                    if (fitPoints.Count < 2) return;

                    // Vẽ polyline fillet màu vàng với bo góc
                    wd.SubEntityTraits.Color = 2; // Vàng
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




            private void DrawTickPreview(Autodesk.AutoCAD.GraphicsInterface.WorldDraw wd)
            {
                try
                {
                    // Tạo polyline fit để vẽ tick trên polyline màu vàng
                    List<Point3d> points = new List<Point3d>();
                    for (int i = 0; i < _polyline.NumberOfVertices; i++)
                    {
                        points.Add(_polyline.GetPoint3dAt(i));
                    }
                    
                    // Thêm điểm preview nếu có
                    if (_hasPreviewPoint)
                    {
                        points.Add(_currentPoint);
                    }

                    if (points.Count < 2) return;

                    // Tạo polyline fit với bo góc giống polyline tạm
                    FilletOverrule overrule = new FilletOverrule();
                    var tempPl = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    
                    // Copy điểm vào polyline tạm
                    for (int i = 0; i < points.Count; i++)
                    {
                        tempPl.AddVertexAt(i, new Point2d(points[i].X, points[i].Y), 0, 0, 0);
                    }
                    
                    List<Point3d> fitPoints = overrule.CreateTempPolylinePoints(tempPl);
                    tempPl.Dispose();

                    if (fitPoints.Count < 2) return;

                    // Tính tổng chiều dài polyline fit
                    double totalLength = 0.0;
                    for (int i = 1; i < fitPoints.Count; i++)
                    {
                        totalLength += (fitPoints[i] - fitPoints[i - 1]).Length;
                    }

                    if (totalLength <= 0.0) return;

                    // Số lượng tick
                    var tickCount = (int)Math.Round(totalLength / TICK_SPACING);
                    if (tickCount < 1) tickCount = 1;
                    var actualSpacing = totalLength / tickCount;

                    // Tham số tick
                    var totalTickLength = RED_LENGTH + _blueWidth + RED_LENGTH;
                    var halfTickLength = totalTickLength / 2.0;

                    // Vẽ tick preview trên polyline fit
                    for (int i = 0; i <= tickCount; i++)
                    {
                        double distance = i * actualSpacing;
                        if (distance > totalLength) distance = totalLength;

                        try
                        {
                            // Tìm điểm trên polyline fit tại khoảng cách distance
                            double walked = 0.0;
                            Point3d pointOnPl = fitPoints[0];

                            for (int j = 1; j < fitPoints.Count; j++)
                            {
                                Point3d p1 = fitPoints[j - 1];
                                Point3d p2 = fitPoints[j];
                                double segLength = (p2 - p1).Length;

                                if (walked + segLength >= distance)
                                {
                                    double t = (distance - walked) / segLength;
                                    pointOnPl = p1 + (p2 - p1) * t;
                                    break;
                                }
                                walked += segLength;
                            }

                        // Tính hướng tiếp tuyến tại điểm hiện tại trên polyline fit
                        Vector3d tangent = new Vector3d(1, 0, 0); // Mặc định

                        // Tìm đoạn chứa điểm hiện tại trên polyline fit
                        double walked2 = 0.0;
                        for (int k = 1; k < fitPoints.Count; k++)
                        {
                            Point3d p1 = fitPoints[k - 1];
                            Point3d p2 = fitPoints[k];
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
                        var tickStart = pointOnPl - normal * halfTickLength;
                        var redEnd1 = pointOnPl - normal * (halfTickLength - RED_LENGTH);
                        var blueStart = pointOnPl - normal * (halfTickLength - RED_LENGTH);
                        var blueEnd = pointOnPl + normal * (halfTickLength - RED_LENGTH);
                        var redStart2 = pointOnPl + normal * (halfTickLength - RED_LENGTH);
                        var tickEnd = pointOnPl + normal * halfTickLength;

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
                catch
                {
                    // Ignore errors in preview
                }
            }
        }
    }
    }
}