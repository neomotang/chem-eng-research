%Script is used to perform nonlinear least squares and absolute deviation regression, fitting hydrodynamic data (liquid holdup, h, and some dimensionless variables, Fr2Re) to a power law expression of the form y = a*x^b

%Representative synthetic data are included below to demonstrate use

h = [0.0191; 0.0806; 0.1501; 0.2125; 0.2747; 0.3506; 0.4017; 0.0514; 0.1087; 0.1231]; %experimental holdup data
Fr2Re = [9.424E-10; 2.988E-09; 4.976E-09; 5.280E-09; 1.954E-08; 1.695E-08; 6.387E-08; 1.336E-09; 2.671E-09; 9.518E-09]; %Froude number^2 * Reynolds number

lnh = [-3.958066944; -2.518256629; -1.89645354; -1.548813291; -1.292075686; -1.048109306; -0.912049738; -2.968117107; -2.219163485; -2.094758246]; %natural logarithm of h
lnFr2Re = [-20.7826086; -19.6286688; -19.11859551; -19.05930921; -17.75086688; -17.89285642; -16.56642797; -20.43323582; -19.74062864; -18.47013363]; %natural logarithm of Fr2Re

x = size(h);

nLS = fit(Fr2Re,h,'power1'); %nonlinear regression using least squares
nLA = fit(Fr2Re,h,'power1',Robust="LAR"); %nonlinear regression using least absolute deviation
lLA = fit(lnFr2Re,lnh,'poly1',Robust="LAR"); %linear regression using least absolute deviation, based on the linearized data

h_LS = nLS.a*Fr2Re.^nLS.b; %predicted holdup from the nonlinear least squares regression
h_LA = nLA.a*Fr2Re.^nLA.b; %predicted holdup from the nonlinear least absolute deviation regression
h_lLA = exp(lLA.p2)*Fr2Re.^lLA.p1; %predicted holdup from the linear least absolute deviation regression

%residuals from the three regression cases
res_LS = h - h_LS;
res_LA = h - h_LA;
res_lLA = h - h_lLA;

%percentage average absolute deviations from the three regression cases
aad_LS = sum(abs(res_LS./h))/x(1,1)*100;
aad_LA = sum(abs(res_LA./h))/x(1,1)*100;
aad_lLA = sum(abs(res_lLA./h))/x(1,1)*100;

%figure,
%plot(Fr2Re,h,'o','MarkerFaceColor','b','MarkerEdgeColor','b')
%hold on
%p = zeros(3,3);

% Plot LS Line
%x_line = -2.5:.1:2.5;
%p(:,2) = plot(Fr2Re,h_LS,'m');
%p(:,2) = plot(Fr2Re,h_LS,'o','MarkerFaceColor','m','MarkerEdgeColor','m');
%set(p(:,2),'LineWidth',1.2)
%hold on

% Plot LAD Line
%x_line = -2.5:.1:2.5;
%p(:,3) = plot(Fr2Re,h_LA,'g');
%p(:,3) = plot(Fr2Re,h_LA,'o','MarkerFaceColor','g','MarkerEdgeColor','g');
%set(p(:,3),'LineWidth',1.2)
%legend(p(1,:), {'Original','LS','LAD'})
%legend({'Original','LS','LAD'})

%figure,
%plot(Fr2Re,h,'o','MarkerFaceColor','b','MarkerEdgeColor','b')
%hold on
%p = zeros(2,2);

% Plot LS residuals Line
%x_line = -2.5:.1:2.5;
%figure,
%plot(Fr2Re,res_LS,'o','MarkerFaceColor','m','MarkerEdgeColor','m');
%set(p(:,2),'LineWidth',1.2)
%hold on

% Plot LAD residuals Line
%x_line = -2.5:.1:2.5;
%p(:,2) = plot(Fr2Re,res_LA,'o','MarkerFaceColor','g','MarkerEdgeColor','g');
%set(p(:,3),'LineWidth',1.2)
%hold on

%p(:,3) = plot(Fr2Re,res_lLA,'o','MarkerFaceColor','b','MarkerEdgeColor','b');

%legend(p(1,:), {'Original','LS','LAD'})
%legend({'LS','LAD','lLAD'})

