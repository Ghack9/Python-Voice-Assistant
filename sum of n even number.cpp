#include <bits/stdc++.h>
using namespace std;
 
// function to find sum of
// first n even numbers
int evenSum(int n)
{
    // required sum
    return (n * (n + 1));
}
 
// Driver program to test above
int main()
{
    int n = 20;
    cout << "Sum of first " << n
         << " Even numbers is: " << evenSum(n);
    return 0;
}
