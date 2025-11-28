%script file for V - volume of reactor as a function of time
function V = volume(t, t1) % t received as a column vector, t1 is the time required to feed the MEA solution
dim = size(t);
V = zeros(dim); % creates an array of zeros, same dimensions as t
for k=1:dim(1) % k - counter used to go through each element in the vector
    if (t(k) <= 0)
        V(k) = 0.500;
    else if ((t(k) <= t1)&&(t(k) > 0))
            V(k) = 0.500 + 0.090/t1*t(k);
        else if (t(k) > t1)
                V(k) = 0.590;
            end
        end
    end
end
end
