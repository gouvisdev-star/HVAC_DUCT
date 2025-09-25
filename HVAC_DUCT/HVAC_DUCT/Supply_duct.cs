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
        /// DrawableOverrule để vẽ thêm các tick vuông góc với Polyline
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
            /// Override method WorldDraw để vẽ thêm các tick vuông góc với polyline và fillet radius tại các góc
            /// </summary>
            /// <param name="drawable">Đối tượng cần vẽ</param>
            /// <param name="wd">WorldDraw context</param>
            /// <returns>True nếu vẽ thành công</returns>
            public override bool WorldDraw(Drawable drawable, WorldDraw wd)
            {
                // Gọi vẽ mặc định trước
                bool result = base.WorldDraw(drawable, wd);

                // Chỉ áp dụng cho Polyline được phép
                if (drawable is Autodesk.AutoCAD.DatabaseServices.Polyline pl)
                {
                    // Kiểm tra polyline có trong danh sách được phép không
                    if (!_allowedPolylines.Contains(pl.ObjectId))
                    {
                        return result;
                    }

                    // Vẽ fillet radius tại các góc của polyline
                    DrawFilletArcs(pl, wd); // Vẽ fillet bằng Arc
                                            // DrawVirtualSpline(pl, wd); // Tắt spline ảo
                                            // Tính tổng chiều dài polyline
                    double totalLength = 0.0;
                    try { totalLength = pl.GetDistanceAtParameter(pl.EndParam); } catch { totalLength = 0.0; }
                    if (totalLength <= 0.0) return result;

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

                    // Vẽ tick tại các vị trí dọc theo polyline
                    for (int i = 0; i <= tickCount; i++)
                    {
                        double distance = i * actualSpacing;
                        if (distance > totalLength) distance = totalLength;

                        try
                        {
                            // Lấy điểm trên polyline tại khoảng cách distance
                            Point3d pointOnPl = pl.GetPointAtDist(distance);

                            // Tính hướng tiếp tuyến tại điểm này
                            double param = pl.GetParameterAtPoint(pointOnPl);
                            double deltaParam = 0.01;
                            double p0 = Math.Max(pl.StartParam, param - deltaParam);
                            double p1 = Math.Min(pl.EndParam, param + deltaParam);

                            Point3d p0_pt = pl.GetPointAtParameter(p0);
                            Point3d p1_pt = pl.GetPointAtParameter(p1);

                            Vector3d tangent = p1_pt - p0_pt;
                            if (tangent.Length < 1e-9) continue;

                            tangent = tangent.GetNormal();

                            // Vector pháp tuyến (vuông góc với tiếp tuyến)
                            Vector3d normal = new Vector3d(-tangent.Y, tangent.X, 0.0);

                            // Tính các điểm của tick
                            Point3d tickStart = pointOnPl - normal * halfTickLength;
                            Point3d redEnd1 = pointOnPl - normal * (halfTickLength - redLength);
                            Point3d blueStart = pointOnPl - normal * (halfTickLength - redLength);
                            Point3d blueEnd = pointOnPl + normal * (halfTickLength - redLength);
                            Point3d redStart2 = pointOnPl + normal * (halfTickLength - redLength);
                            Point3d tickEnd = pointOnPl + normal * halfTickLength;

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
                return result;
            }

            /// <summary>
            /// Vẽ fillet bằng Arc tại các góc của polyline
            /// </summary>
            /// <param name="pl">Polyline cần vẽ fillet</param>
            /// <param name="wd">WorldDraw context</param>
            private void DrawFilletArcs(Autodesk.AutoCAD.DatabaseServices.Polyline pl, WorldDraw wd)
            {
                if (pl.NumberOfVertices < 3) return;

                double filletRadius = 8.0; // Bán kính fillet mặc định

                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    // Lấy 3 đỉnh liên tiếp
                    int prev = (i - 1 + pl.NumberOfVertices) % pl.NumberOfVertices;
                    int curr = i;
                    int next = (i + 1) % pl.NumberOfVertices;

                    Point3d p1 = pl.GetPoint3dAt(prev);
                    Point3d p2 = pl.GetPoint3dAt(curr);
                    Point3d p3 = pl.GetPoint3dAt(next);

                    // Vẽ fillet arc tại góc
                    DrawFilletArcAtVertex(wd, p1, p2, p3, filletRadius);
                }
            }

            /// <summary>
            /// Vẽ fillet arc tại một đỉnh cụ thể
            /// </summary>
            /// <param name="wd">WorldDraw context</param>
            /// <param name="p1">Điểm trước</param>
            /// <param name="p2">Điểm góc</param>
            /// <param name="p3">Điểm sau</param>
            /// <param name="radius">Bán kính fillet</param>
            private void DrawFilletArcAtVertex(WorldDraw wd, Point3d p1, Point3d p2, Point3d p3, double radius)
            {
                try
                {
                    // Tính vector từ góc ra ngoài
                    Vector3d v1 = (p2 - p1).GetNormal(); // Từ p1 đến p2
                    Vector3d v2 = (p2 - p3).GetNormal(); // Từ p3 đến p2

                    // Tính góc giữa hai vector
                    double cosAngle = v1.DotProduct(v2);
                    cosAngle = Math.Max(-1.0, Math.Min(1.0, cosAngle));
                    double angle = Math.Acos(cosAngle);

                    // Nếu góc quá nhỏ hoặc quá lớn, bỏ qua
                    if (angle < 0.1 || angle > Math.PI - 0.1) return;

                    // Tính khoảng cách từ góc đến điểm bắt đầu fillet
                    double distance = radius / Math.Tan(angle / 2.0);

                    // Tính điểm bắt đầu và kết thúc của fillet
                    Point3d startPoint = p2 + v1 * distance;
                    Point3d endPoint = p2 + v2 * distance;

                    // Tính tâm cung fillet
                    Vector3d bisector = (v1 + v2).GetNormal();
                    double centerDistance = radius / Math.Sin(angle / 2.0);
                    Point3d center = p2 + bisector * centerDistance;

                    // Vẽ cung fillet bằng nhiều đoạn thẳng
                    DrawArcSegments(wd, center, startPoint, endPoint, radius);
                }
                catch
                {
                    // Bỏ qua lỗi nếu có
                }
            }

            /// <summary>
            /// Vẽ cung tròn bằng nhiều đoạn thẳng
            /// </summary>
            /// <param name="wd">WorldDraw context</param>
            /// <param name="center">Tâm cung</param>
            /// <param name="startPoint">Điểm bắt đầu</param>
            /// <param name="endPoint">Điểm kết thúc</param>
            /// <param name="radius">Bán kính</param>
            private void DrawArcSegments(WorldDraw wd, Point3d center, Point3d startPoint, Point3d endPoint, double radius)
            {
                // Tính góc bắt đầu và kết thúc
                Vector3d vStart = (startPoint - center).GetNormal();
                Vector3d vEnd = (endPoint - center).GetNormal();

                double startAngle = Math.Atan2(vStart.Y, vStart.X);
                double endAngle = Math.Atan2(vEnd.Y, vEnd.X);

                // Điều chỉnh góc nếu cần
                if (endAngle < startAngle)
                {
                    endAngle += 2 * Math.PI;
                }

                // Vẽ cung bằng nhiều đoạn thẳng
                int segments = 12; // Số đoạn để tạo cung mượt
                for (int i = 0; i < segments; i++)
                {
                    double t1 = (double)i / segments;
                    double t2 = (double)(i + 1) / segments;

                    double angle1 = startAngle + (endAngle - startAngle) * t1;
                    double angle2 = startAngle + (endAngle - startAngle) * t2;

                    Point3d pt1 = center + new Vector3d(Math.Cos(angle1), Math.Sin(angle1), 0) * radius;
                    Point3d pt2 = center + new Vector3d(Math.Cos(angle2), Math.Sin(angle2), 0) * radius;

                    wd.Geometry.WorldLine(pt1, pt2);
                }
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

            /// <summary>
            /// Lệnh tạo fillet với Arc + Wipeout cho polyline hiện có
            /// </summary>
            [CommandMethod("FILLET_ARC_WIPEOUT")]
            public static void FilletArcWipeout()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Chọn polyline
                    PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline để tạo fillet: ");
                    peo.SetRejectMessage("\nVui lòng chọn polyline!");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    // Nhập bán kính fillet
                    PromptDoubleOptions radiusOpts = new PromptDoubleOptions("\nNhập bán kính fillet: ");
                    radiusOpts.AllowNegative = false;
                    radiusOpts.AllowZero = false;
                    radiusOpts.DefaultValue = 8.0;

                    PromptDoubleResult radiusResult = ed.GetDouble(radiusOpts);
                    if (radiusResult.Status != PromptStatus.OK) return;

                    double filletRadius = radiusResult.Value;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Polyline pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;

                        if (pl != null)
                        {
                            // Tạo fillet thực tế cho polyline
                            CreateRealFilletArcs(pl, filletRadius, tr);

                            tr.Commit();
                            ed.WriteMessage($"\nĐã tạo fillet với bán kính {filletRadius:F1}!");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Tạo fillet thực tế cho polyline bằng Arc + Wipeout
            /// </summary>
            private static void CreateRealFilletArcs(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double radius, Transaction tr)
            {
                try
                {
                    // Kiểm tra polyline hợp lệ
                    if (pl == null || pl.IsErased || pl.IsDisposed)
                        return;

                    BlockTableRecord btr = tr.GetObject(pl.BlockId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr == null)
                        return;

                    if (pl.NumberOfVertices < 3) return;

                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        // Lấy 3 đỉnh liên tiếp
                        int prev = (i - 1 + pl.NumberOfVertices) % pl.NumberOfVertices;
                        int curr = i;
                        int next = (i + 1) % pl.NumberOfVertices;

                        Point3d p1 = pl.GetPoint3dAt(prev);
                        Point3d p2 = pl.GetPoint3dAt(curr);
                        Point3d p3 = pl.GetPoint3dAt(next);

                        // Tạo fillet arc tại góc
                        CreateFilletArcAtVertexSimple(btr, p1, p2, p3, radius, tr);
                    }
                }
                catch (System.Exception ex)
                {
                    // Bỏ qua lỗi nếu có
                }
            }


            /// <summary>
            /// Tạo fillet arc thực tế tại một đỉnh (phiên bản đơn giản)
            /// </summary>
            private static void CreateFilletArcAtVertexSimple(BlockTableRecord btr, Point3d p1, Point3d p2, Point3d p3, double radius, Transaction tr)
            {
                // Gọi method bool và bỏ qua kết quả trả về
                CreateFilletArcAtVertex(btr, p1, p2, p3, radius, tr);
            }

            /// <summary>
            /// Tạo Wipeout để che khuất phần giao nhau của fillet
            /// </summary>
            private static void CreateWipeoutForFillet(BlockTableRecord btr, Point3d vertex, Point3d startPoint, Point3d endPoint, Point3d center, double radius, Transaction tr)
            {
                try
                {
                    // Kiểm tra tính hợp lệ của các điểm
                    if (vertex == Point3d.Origin || startPoint == Point3d.Origin || endPoint == Point3d.Origin)
                        return;

                    // Kiểm tra khoảng cách tối thiểu giữa các điểm
                    double minDistance = 0.001;
                    if ((startPoint - vertex).Length < minDistance || 
                        (endPoint - vertex).Length < minDistance || 
                        (endPoint - startPoint).Length < minDistance)
                        return;

                    // Tạo boundary cho wipeout (hình tam giác)
                    Point2dCollection boundaryPoints = new Point2dCollection();
                    boundaryPoints.Add(new Point2d(vertex.X, vertex.Y));
                    boundaryPoints.Add(new Point2d(startPoint.X, startPoint.Y));
                    boundaryPoints.Add(new Point2d(endPoint.X, endPoint.Y));

                    // Kiểm tra boundary có ít nhất 3 điểm không
                    if (boundaryPoints.Count < 3)
                        return;

                    // Tạo wipeout với validation
                    Wipeout wipeout = new Wipeout();
                    wipeout.SetDatabaseDefaults();
                    
                    // Sử dụng try-catch riêng cho SetFrom
                    try
                    {
                        wipeout.SetFrom(boundaryPoints, Vector3d.ZAxis);
                        
                        // Kiểm tra wipeout có hợp lệ không
                        if (wipeout != null && !wipeout.IsErased)
                        {
                            btr.AppendEntity(wipeout);
                            tr.AddNewlyCreatedDBObject(wipeout, true);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // Dispose wipeout nếu có lỗi
                        if (wipeout != null)
                        {
                            wipeout.Dispose();
                        }
                        // Log lỗi nếu cần (có thể bỏ qua)
                    }
                }
                catch (System.Exception ex)
                {
                    // Bỏ qua lỗi nếu có
                }
            }

            /// <summary>
            /// Lệnh demo fillet với preview real-time
            /// </summary>
            [CommandMethod("FILLET_PREVIEW")]
            public static void FilletPreview()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Chọn polyline
                    PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline để preview fillet: ");
                    peo.SetRejectMessage("\nVui lòng chọn polyline!");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    // Nhập bán kính fillet
                    PromptDoubleOptions radiusOpts = new PromptDoubleOptions("\nNhập bán kính fillet: ");
                    radiusOpts.AllowNegative = false;
                    radiusOpts.AllowZero = false;
                    radiusOpts.DefaultValue = 8.0;

                    PromptDoubleResult radiusResult = ed.GetDouble(radiusOpts);
                    if (radiusResult.Status != PromptStatus.OK) return;

                    double filletRadius = radiusResult.Value;

                    // Bật overrule với fillet radius
                    if (_overrule == null)
                    {
                        _overrule = new FilletOverrule();
                        Overrule.AddOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule, true);
                        Overrule.Overruling = true;
                    }

                    // Thêm polyline vào danh sách được phép hiển thị fillet
                    FilletOverrule.AddAllowedPolyline(per.ObjectId);

                    ed.WriteMessage($"\nĐã bật preview fillet với bán kính {filletRadius:F1}! Gõ FILLET_OFF để tắt.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Lệnh tắt fillet overrule
            /// </summary>
            [CommandMethod("FILLET_OFF")]
            public static void FilletOff()
            {
                if (_overrule != null)
                {
                    Overrule.RemoveOverrule(RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)), _overrule);
                    _overrule = null;
                    Overrule.Overruling = false;

                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Editor ed = doc.Editor;
                    ed.WriteMessage("\nĐã tắt fillet overrule!");
                }
            }

            /// <summary>
            /// Lệnh xóa bỏ tất cả fillet arc khỏi polyline
            /// </summary>
            [CommandMethod("REMOVE_FILLET")]
            public static void RemoveFillet()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Chọn polyline
                    PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline để xóa fillet: ");
                    peo.SetRejectMessage("\nVui lòng chọn polyline!");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Polyline pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;

                        if (pl != null)
                        {
                            // Xóa tất cả fillet arc xung quanh polyline
                            int removedCount = RemoveAllFilletArcs(pl, tr);

                            tr.Commit();
                            ed.WriteMessage($"\nĐã xóa {removedCount} fillet arc!");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Xóa tất cả fillet arc xung quanh polyline
            /// </summary>
            private static int RemoveAllFilletArcs(Autodesk.AutoCAD.DatabaseServices.Polyline pl, Transaction tr)
            {
                int removedCount = 0;
                BlockTableRecord btr = tr.GetObject(pl.BlockId, OpenMode.ForRead) as BlockTableRecord;
                
                if (btr == null) return 0;

                try
                {
                    // Lấy tất cả entities trong block
                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (ent != null && ent is Arc)
                        {
                            Arc arc = ent as Arc;
                            
                            // Kiểm tra xem arc có phải là fillet arc không
                            if (IsFilletArc(arc, pl))
                            {
                                // Xóa arc
                                ent.UpgradeOpen();
                                ent.Erase();
                                removedCount++;
                            }
                        }
                        else if (ent != null && ent is Wipeout)
                        {
                            // Xóa wipeout liên quan đến fillet
                            Wipeout wipeout = ent as Wipeout;
                            if (IsFilletWipeout(wipeout, pl))
                            {
                                wipeout.UpgradeOpen();
                                wipeout.Erase();
                                removedCount++;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    // Bỏ qua lỗi nếu có
                }

                return removedCount;
            }

            /// <summary>
            /// Kiểm tra xem arc có phải là fillet arc không
            /// </summary>
            private static bool IsFilletArc(Arc arc, Autodesk.AutoCAD.DatabaseServices.Polyline pl)
            {
                try
                {
                    // Kiểm tra màu sắc (fillets thường có màu xanh lá - ACI 3)
                    if (arc.Color.ColorIndex != 3) return false;

                    // Kiểm tra radius hợp lý (fillets thường có radius từ 1-50)
                    if (arc.Radius < 1.0 || arc.Radius > 50.0) return false;

                    // Kiểm tra xem arc có gần polyline không
                    Point3d arcCenter = arc.Center;
                    double minDistance = double.MaxValue;

                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        Point3d vertex = pl.GetPoint3dAt(i);
                        double distance = arcCenter.DistanceTo(vertex);
                        if (distance < minDistance)
                            minDistance = distance;
                    }

                    // Nếu arc gần polyline (trong vòng 100 units) thì có thể là fillet
                    return minDistance < 100.0;
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// Kiểm tra xem wipeout có phải là fillet wipeout không
            /// </summary>
            private static bool IsFilletWipeout(Wipeout wipeout, Autodesk.AutoCAD.DatabaseServices.Polyline pl)
            {
                try
                {
                    // Kiểm tra xem wipeout có gần polyline không
                    // (Có thể cần logic phức tạp hơn tùy thuộc vào cách tạo wipeout)
                    return true; // Tạm thời xóa tất cả wipeout
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// Lệnh xóa tất cả fillet arc trong toàn bộ drawing
            /// </summary>
            [CommandMethod("REMOVE_ALL_FILLETS")]
            public static void RemoveAllFillets()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        int removedCount = 0;

                        // Tạo danh sách các entity cần xóa
                        List<ObjectId> entitiesToRemove = new List<ObjectId>();

                        foreach (ObjectId objId in btr)
                        {
                            Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                            
                            if (ent != null && ent is Arc)
                            {
                                Arc arc = ent as Arc;
                                
                                // Kiểm tra xem arc có phải là fillet arc không
                                if (IsFilletArcByProperties(arc))
                                {
                                    entitiesToRemove.Add(objId);
                                }
                            }
                            else if (ent != null && ent is Wipeout)
                            {
                                // Xóa tất cả wipeout (có thể liên quan đến fillet)
                                entitiesToRemove.Add(objId);
                            }
                        }

                        // Xóa các entity
                        foreach (ObjectId objId in entitiesToRemove)
                        {
                            Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                            if (ent != null)
                            {
                                ent.Erase();
                                removedCount++;
                            }
                        }

                        tr.Commit();
                        ed.WriteMessage($"\nĐã xóa {removedCount} fillet arc và wipeout!");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Kiểm tra arc có phải là fillet arc dựa trên thuộc tính
            /// </summary>
            private static bool IsFilletArcByProperties(Arc arc)
            {
                try
                {
                    // Kiểm tra màu sắc (fillets thường có màu xanh lá - ACI 3)
                    if (arc.Color.ColorIndex != 3) return false;

                    // Kiểm tra radius hợp lý (fillets thường có radius từ 1-50)
                    if (arc.Radius < 1.0 || arc.Radius > 50.0) return false;

                    // Kiểm tra góc hợp lý (fillets thường có góc nhỏ hơn 180 độ)
                    double angleSpan = Math.Abs(arc.EndAngle - arc.StartAngle);
                    if (angleSpan > Math.PI) return false;

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// Lệnh tự động bo góc nhọn của Polyline
            /// </summary>
            [CommandMethod("AUTO_FILLET_SHARP")]
            public static void AutoFilletSharp()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Chọn polyline
                    PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline để tự động bo góc nhọn: ");
                    peo.SetRejectMessage("\nVui lòng chọn polyline!");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    // Nhập bán kính fillet
                    PromptDoubleOptions radiusOpts = new PromptDoubleOptions("\nNhập bán kính fillet: ");
                    radiusOpts.AllowNegative = false;
                    radiusOpts.AllowZero = false;
                    radiusOpts.DefaultValue = 5.0;

                    PromptDoubleResult radiusResult = ed.GetDouble(radiusOpts);
                    if (radiusResult.Status != PromptStatus.OK) return;

                    double filletRadius = radiusResult.Value;

                    // Nhập góc tối thiểu để bo (độ)
                    PromptDoubleOptions angleOpts = new PromptDoubleOptions("\nNhập góc tối thiểu để bo (độ, mặc định 90): ");
                    angleOpts.AllowNegative = false;
                    angleOpts.AllowZero = false;
                    angleOpts.DefaultValue = 90.0;

                    PromptDoubleResult angleResult = ed.GetDouble(angleOpts);
                    if (angleResult.Status != PromptStatus.OK) return;

                    double minAngleDegrees = angleResult.Value;
                    double minAngleRadians = minAngleDegrees * Math.PI / 180.0;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Polyline pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;

                        if (pl != null)
                        {
                            // Tự động bo các góc nhọn
                            int filletCount = AutoFilletSharpAngles(pl, filletRadius, minAngleRadians, tr);

                            tr.Commit();
                            ed.WriteMessage($"\nĐã bo {filletCount} góc nhọn với bán kính {filletRadius:F1}!");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Tự động bo các góc nhọn của polyline
            /// </summary>
            private static int AutoFilletSharpAngles(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double radius, double minAngle, Transaction tr)
            {
                BlockTableRecord btr = tr.GetObject(pl.BlockId, OpenMode.ForWrite) as BlockTableRecord;
                int filletCount = 0;

                if (pl.NumberOfVertices < 3) return 0;

                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    // Lấy 3 đỉnh liên tiếp
                    int prev = (i - 1 + pl.NumberOfVertices) % pl.NumberOfVertices;
                    int curr = i;
                    int next = (i + 1) % pl.NumberOfVertices;

                    Point3d p1 = pl.GetPoint3dAt(prev);
                    Point3d p2 = pl.GetPoint3dAt(curr);
                    Point3d p3 = pl.GetPoint3dAt(next);

                    // Tính góc tại đỉnh
                    Vector3d v1 = (p2 - p1).GetNormal();
                    Vector3d v2 = (p2 - p3).GetNormal();

                    double cosAngle = v1.DotProduct(v2);
                    cosAngle = Math.Max(-1.0, Math.Min(1.0, cosAngle));
                    double angle = Math.Acos(cosAngle);

                    // Chỉ bo góc nhọn (góc < minAngle)
                    if (angle < minAngle)
                    {
                        // Tạo fillet arc tại góc nhọn
                        if (CreateFilletArcAtVertex(btr, p1, p2, p3, radius, tr))
                        {
                            filletCount++;
                        }
                    }
                }

                return filletCount;
            }

            /// <summary>
            /// Tạo fillet arc thực tế tại một đỉnh (trả về true nếu thành công)
            /// </summary>
            private static bool CreateFilletArcAtVertex(BlockTableRecord btr, Point3d p1, Point3d p2, Point3d p3, double radius, Transaction tr)
            {
                try
                {
                    // Kiểm tra tính hợp lệ của các điểm
                    if (p1 == Point3d.Origin || p2 == Point3d.Origin || p3 == Point3d.Origin)
                        return false;

                    // Kiểm tra radius hợp lệ
                    if (radius <= 0 || radius > 1000)
                        return false;
                    // Tính vector từ góc ra ngoài
                    Vector3d v1 = (p2 - p1).GetNormal();
                    Vector3d v2 = (p2 - p3).GetNormal();

                    // Tính góc giữa hai vector
                    double cosAngle = v1.DotProduct(v2);
                    cosAngle = Math.Max(-1.0, Math.Min(1.0, cosAngle));
                    double angle = Math.Acos(cosAngle);

                    // Nếu góc quá nhỏ hoặc quá lớn, bỏ qua
                    if (angle < 0.1 || angle > Math.PI - 0.1) return false;

                    // Tính khoảng cách từ góc đến điểm bắt đầu fillet
                    double distance = radius / Math.Tan(angle / 2.0);

                    // Kiểm tra khoảng cách có hợp lý không
                    if (distance <= 0 || distance > 1000) return false;

                    // Tính điểm bắt đầu và kết thúc của fillet
                    Point3d startPoint = p2 + v1 * distance;
                    Point3d endPoint = p2 + v2 * distance;

                    // Tính tâm cung fillet
                    Vector3d bisector = (v1 + v2).GetNormal();
                    double centerDistance = radius / Math.Sin(angle / 2.0);
                    Point3d center = p2 + bisector * centerDistance;

                    // Tạo Arc thực tế
                    Arc arc = new Arc();
                    arc.SetDatabaseDefaults();
                    arc.Center = center;
                    arc.Radius = radius;
                    arc.Normal = Vector3d.ZAxis;

                    // Tính góc bắt đầu và kết thúc
                    Vector3d vStart = (startPoint - center).GetNormal();
                    Vector3d vEnd = (endPoint - center).GetNormal();

                    double startAngle = Math.Atan2(vStart.Y, vStart.X);
                    double endAngle = Math.Atan2(vEnd.Y, vEnd.X);

                    // Điều chỉnh góc nếu cần
                    if (endAngle < startAngle)
                    {
                        endAngle += 2 * Math.PI;
                    }

                    arc.StartAngle = startAngle;
                    arc.EndAngle = endAngle;
                    arc.Color = Color.FromColorIndex(ColorMethod.ByAci, 3); // Xanh lá

                    btr.AppendEntity(arc);
                    tr.AddNewlyCreatedDBObject(arc, true);

                    // Tạm thời vô hiệu hóa Wipeout để tránh crash
                    // TODO: Kích hoạt lại khi đã fix lỗi Wipeout
                    // CreateWipeoutForFillet(btr, p2, startPoint, endPoint, center, radius, tr);

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// Lệnh bo tất cả góc của Polyline
            /// </summary>
            [CommandMethod("FILLET_ALL_ANGLES")]
            public static void FilletAllAngles()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                try
                {
                    // Chọn polyline
                    PromptEntityOptions peo = new PromptEntityOptions("\nChọn polyline để bo tất cả góc: ");
                    peo.SetRejectMessage("\nVui lòng chọn polyline!");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);

                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    // Nhập bán kính fillet
                    PromptDoubleOptions radiusOpts = new PromptDoubleOptions("\nNhập bán kính fillet: ");
                    radiusOpts.AllowNegative = false;
                    radiusOpts.AllowZero = false;
                    radiusOpts.DefaultValue = 3.0;

                    PromptDoubleResult radiusResult = ed.GetDouble(radiusOpts);
                    if (radiusResult.Status != PromptStatus.OK) return;

                    double filletRadius = radiusResult.Value;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.Polyline pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;

                        if (pl != null)
                        {
                            // Bo tất cả góc
                            int filletCount = FilletAllAngles(pl, filletRadius, tr);

                            tr.Commit();
                            ed.WriteMessage($"\nĐã bo {filletCount} góc với bán kính {filletRadius:F1}!");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nLỗi: {ex.Message}");
                }
            }

            /// <summary>
            /// Bo tất cả góc của polyline
            /// </summary>
            private static int FilletAllAngles(Autodesk.AutoCAD.DatabaseServices.Polyline pl, double radius, Transaction tr)
            {
                BlockTableRecord btr = tr.GetObject(pl.BlockId, OpenMode.ForWrite) as BlockTableRecord;
                int filletCount = 0;

                if (pl.NumberOfVertices < 3) return 0;

                for (int i = 0; i < pl.NumberOfVertices; i++)
                {
                    // Lấy 3 đỉnh liên tiếp
                    int prev = (i - 1 + pl.NumberOfVertices) % pl.NumberOfVertices;
                    int curr = i;
                    int next = (i + 1) % pl.NumberOfVertices;

                    Point3d p1 = pl.GetPoint3dAt(prev);
                    Point3d p2 = pl.GetPoint3dAt(curr);
                    Point3d p3 = pl.GetPoint3dAt(next);

                    // Bo tất cả góc (trừ góc quá nhỏ)
                    if (CreateFilletArcAtVertex(btr, p1, p2, p3, radius, tr))
                    {
                        filletCount++;
                    }
                }

                return filletCount;
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
                // Vẽ polyline
                if (_polyline.NumberOfVertices >= 2)
                {
                    for (int i = 1; i < _polyline.NumberOfVertices; i++)
                    {
                        Point3d p1 = _polyline.GetPoint3dAt(i - 1);
                        Point3d p2 = _polyline.GetPoint3dAt(i);
                        wd.Geometry.WorldLine(p1, p2);
                    }
                }

                // Vẽ tick preview
                if (_polyline.NumberOfVertices >= 2)
                {
                    DrawTickPreview(wd);
                }

                return true;
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
