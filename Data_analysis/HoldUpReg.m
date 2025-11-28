clear;

h = [0.060616683; 0.068904886; 0.138016332; 0.169026769; 0.287905755; 0.274849378; 0.490112363; 0.058988007; 0.078305391; 0.161587108; 0.114772076; 0.156745006; 0.115280419; 0.108966694];
Fr2Re = [7.3053E-10; 3.01816E-09; 6.72462E-09; 6.36164E-09; 2.32604E-08; 2.3545E-08; 6.7946E-08; 1.17234E-09; 3.42499E-09; 1.0125E-08; 1.50692E-09; 4.8193E-09; 1.38576E-09; 1.43812E-09];

x = size(h);

%lnh = [-2.803185127; -2.675028192; -1.980383253; -1.777698178; -1.24512209; -1.291532046; -0.713120603; -2.830421125; -2.547138832; -1.822710914; -2.164807062; -1.853134959; -2.160387692; -2.216713004];
%lnFr2Re = [-21.03725107; -19.61861799; -18.81749048; -18.87298007; -17.57651194; -17.56435285; -16.50455203; -20.56426377; -19.49216803; -18.40826074; -20.31319603; -19.15063699; -20.39701497; -20.35993121];

nLS = fit(Fr2Re,h,'power1');
nLA = fit(Fr2Re,h,'power1',Robust="LAR");
%lLA = fit(lnFr2Re,lnh,'poly1',Robust="LAR");

h_LS = nLS.a*Fr2Re.^nLS.b;
h_LA = nLA.a*Fr2Re.^nLA.b;
%h_lLA = exp(lLA.p2)*Fr2Re.^lLA.p1;

res_LS = h - h_LS;
res_LA = h - h_LA;
%res_lLA = h - h_lLA;

aad_LS = sum(abs(res_LS./h))/x(1,1)*100; % for %AAD
aad_LA = sum(abs(res_LA./h))/x(1,1)*100;
%aad_lLA = sum(abs(res_lLA./h))/x(1,1)*100;

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
