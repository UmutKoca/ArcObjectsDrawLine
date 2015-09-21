 private void btnCalculate_Click(object sender, EventArgs e)
        {

            // Create a polyline 

            IPoint p1 = new PointClass();
            p1.X = Convert.ToDouble(txtP1x.Text); p1.Y = Convert.ToDouble(txtP1y.Text);

            IPoint p2 = new PointClass();
            p2.X = Convert.ToDouble(txtP2x.Text) ; p2.Y = Convert.ToDouble(txtP2y.Text);

            IPoint p3 = new PointClass();
            p3.X = Convert.ToDouble(txtP3x.Text) ; p3.Y = Convert.ToDouble(txtP3y.Text);

            IPoint p4 = new PointClass();
            p4.X = Convert.ToDouble(txtP4x.Text); p4.Y = Convert.ToDouble(txtP4y.Text);

            IPoint p5 = new PointClass();
            p5.X = Convert.ToDouble(txtP5x.Text); p5.Y = Convert.ToDouble(txtP5y.Text);

            IPoint p6 = new PointClass();
            p6.X = Convert.ToDouble(txtP6x.Text); p6.Y = Convert.ToDouble(txtP6y.Text);
       

            IPolyline polyline = new PolylineClass();


            IPointCollection pointColl = polyline as IPointCollection;

            pointColl.AddPoint(p1);
            pointColl.AddPoint(p2);
            pointColl.AddPoint(p3);
            pointColl.AddPoint(p4);
            pointColl.AddPoint(p5);
            pointColl.AddPoint(p6);

            // Create a colour 

            IRgbColor color = new RgbColorClass();
            color.Red = 0; color.Blue = 0; color.Green = 0;

            // Create a line symbol

            ISimpleLineSymbol simpleLineSymbol = new SimpleLineSymbolClass();
            simpleLineSymbol.Color = color;
            simpleLineSymbol.Width = 1;

            // Create a line element, this is the graphic that will get added to container

            IElement element = new LineElement();
            element.Geometry = polyline;

            ILineElement lineElement;
            lineElement = element as ILineElement;
            lineElement.Symbol = simpleLineSymbol;

            // Get the Mxd

            IMxDocument mxdoc = ArcMap.Application.Document as IMxDocument;

            // Get the handle on the page layout

            IPageLayout pageLayout = new PageLayout();
            pageLayout = mxdoc.PageLayout;

            // Get the graphics container of the PageLayout

            IGraphicsContainer graphicsContainer = pageLayout as IGraphicsContainer;

            // Add element an refresh
            graphicsContainer.AddElement(element, 0);

            IActiveView activeView = pageLayout as IActiveView;
            
            activeView.Refresh();

            this.Close();

            
            
        }
