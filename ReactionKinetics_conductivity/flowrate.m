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

