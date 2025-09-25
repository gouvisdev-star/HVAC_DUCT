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
            /// Override method WorldDraw để vẽ thêm các tick vuông góc với polyline (không có fillet arc)
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

                    // Tắt vẽ fillet radius tại các góc của polyline
                    // DrawFilletArcs(pl, wd); // Vẽ fillet bằng Arc - ĐÃ TẮT
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