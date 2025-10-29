function out = MEAfittingzwit(expdata)
%%Rate constant estimation
% Inputs:
% Experimental data (concentration-time profile): expdata
% Concentration equation (dC/dt = f_C(r)) [or conversion X, in this case]
% Rate equation (r = f_r(k))
% Lower and upper bounds for k

% Bounds for k
%%bounds = [6e-08 6e-06]; {Simplifying program}%

% Relevant time span
tspan = expdata(:,1);

% Experimental conversions
Xexp = expdata(:,2);

% Bounds and initial guess for k
%%lb = bounds(1);
%%ub = bounds(2);
k0 = [1]; %k0(1) = initial guess for rate constant;

% Wrapper function for minimization function
PredictionErrork = @(k) PredictionError(Xexp, ConversionPrediction(tspan, k));

% Minimization
options = optimset('largescale', 'off', 'display', 'off');
%[out.kOpt, out.predErrOpt, out.exitflag, out.output] = fminsearch(PredictionErrork, k0, options);
%[out.kOpt, out.predErrOpt, out.exitflag, out.output] = fmincon(PredictionErrork,k0,[],[],[],[],zeros(1,3),20*ones(1,3)[],options);
[out.kOpt, out.predErrOpt, out.exitflag, out.output] = fminsearchbnd(PredictionErrork, k0, zeros(1,1),inf*ones(1,1),options);
% Calculate profile for optimum k
% Wrapper to pass constant k value
ConversionEquationConstantk = @(t,X) ConversionEquation(t,X,out.kOpt);
[tpredOpt,XpredOpt] = ode15s(ConversionEquationConstantk,[tspan(1) tspan(end)],0);

% Plots
figure;
% Predicted profile
plot(tpredOpt, XpredOpt, 'k-');
hold on;
% Actual data
plot(tspan, Xexp, 'rx');
hold off

legend('Predicted profile', 'Actual data');

end

function dX = ConversionEquation(t, X, k)
% Change of conversion with time

% Constant parameters
Na0 = 0.002681209; % initial mol CO2
R = 1.009467872; % ratio of water to MEA in deprotonation reaction
t1 = 142.27/80*90; % time to add MEA solution to the reactor
mol = 0.999052399; % molarity of MEA solution

% Known system variables
V = volume(t, t1); % system volume as a function of time
Fm = flowrate(t, t1, mol); % MEA flow rate as a function of time

% Differential equation
dX = V*((R+1)*(1-X)*Fm*t-(R+2)*(1-X)*X*Na0)/(k(1)*(R+1)*V^2);

end

function Xpred = ConversionPrediction(tspan, k)
% Predicting the conversion values at specific time points (tspan),
% for a given choice of the rate constant k

% Initial conditions for conversion
X0 = 0;

% Wrapper to pass ocnstant k value
ConversionEquationConstantk = @(t,X) ConversionEquation(t,X,k);

% Ordinary differential equation solving
[tpred, Xpred] = ode15s(ConversionEquationConstantk,tspan,X0);

end

function predErr = PredictionError(Xexp, Xpred)
% Sum of squared errors for prediction associated with a certain k value

predErr = sum((Xexp-Xpred).^2);

end

%script file for Fm - Flow rate of MEA as a function of time
function Fm = flowrate(t, t1, mol) % t received as a column vector, t1 is the time to feed the MEA solution, mol is the molarity of the MEA solution [mol/L]
dim = size(t);
Fm = zeros(dim); % creates an array of zeros, same dimensions as t
for k=1:dim(1) % k - counter used to go through each element in the vector
    if ((t(k) <= t1)&&(t(k) >= 0))
        Fm(k) = mol*0.090/t1;
    else
        Fm(k) = 0;
    end
end
end
