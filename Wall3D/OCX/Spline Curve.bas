Attribute VB_Name = "SplineCurve"
Option Explicit
'#define N 3
'XYZ inp[N+1] = {0.0,0.0,0.0,   1.0,0.0,3.0,   2.0,0.0,1.0,   4.0,0.0,4.0};
'#define T 3
'int knots[N+T+1];
'#define RESOLUTION 200
'XYZ outp[RESOLUTION];
Private Const N% = 3
Private inp(N% + 1) As D3DVECTOR
Private Const T% = 3
Private knots%(N% + T% + 1)
Private Const RESOLUTION% = 200
Private outp(RESOLUTION) As D3DVECTOR

'
' This returns the point "output" on the spline curve.
' The parameter "v" indicates the position, it ranges from 0 to n-t+2
'
Sub SplinePoint(u() As Integer, N%, T%, v&, Control() As D3DVECTOR, output As D3DVECTOR)
'int *u,n,t;
'double v;
'XYZ *control,*output;
'{
'   int k;
'   double b;

'   output->x = 0;
'   output->y = 0;
'   output->z = 0;

'   for (k=0;k<=n;k++) {
'      b = SplineBlend(k,t,u,v);
'      output->x += control[k].x * b;
'      output->y += control[k].y * b;
'      output->z += control[k].z * b;
'   }
    
    Dim k%
    Dim b&
    
    output.X = 0
    output.Y = 0
    output.z = 0
    
    For k% = 0 To N% + 1
        b& = SplineBlend(k%, T%, u%(), v&)
        output.X = output.X + Control(k).X * b&
        output.Y = output.Y + Control(k).Y * b&
        output.z = output.z + Control(k).z * b&
    Next k%
End Sub

'/*
'   Calculate the blending value, this is done recursively.
'
'   If the numerator and denominator are 0 the expression is 0.
'   If the deonimator is 0 the expression is 0
'*/
Function SplineBlend(k%, T%, u() As Integer, v&) As Long
'int k,t,*u;
'double v;
'{
'   double value;

'   if (t == 1) {
'      if ((u[k] <= v) && (v < u[k+1]))
'         value = 1;
'      Else
'         value = 0;
'   } else {
'      if ((u[k+t-1] == u[k]) && (u[k+t] == u[k+1]))
'         value = 0;
'      else if (u[k+t-1] == u[k])
'         value = (u[k+t] - v) / (u[k+t] - u[k+1]) * SplineBlend(k+1,t-1,u,v);
'      else if (u[k+t] == u[k+1])
'         value = (v - u[k]) / (u[k+t-1] - u[k]) * SplineBlend(k,t-1,u,v);
'     Else
'         value = (v - u[k]) / (u[k+t-1] - u[k]) * SplineBlend(k,t-1,u,v) +
'                 (u[k+t] - v) / (u[k+t] - u[k+1]) * SplineBlend(k+1,t-1,u,v);
'   }
'   return(value);

    Dim value&
    If T% = 1 Then
        If u%(k%) <= v& And v& < u%(k% + 1) Then
            value& = 1
        Else
            value& = 0
        End If
    Else
        If u%(k% + T% - 1) = u%(k%) And u%(k% + T%) = u%(k% + 1) Then
            value& = 0
        ElseIf u%(k% + T% - 1) = u%(k%) Then
            value& = (u%(k% + T%) - v&) / (u%(k% + T%) - u%(k% + 1)) * SplineBlend(k% + 1, T% - 1, u%(), v&)
        ElseIf u%(k% + T%) = u%(k% + 1) Then
            value& = (v& - u%(k%)) / (u%(k% + T% - 1) - u%(k%)) * SplineBlend(k%, T% - 1, u%(), v&)
        Else
            value& = (v& - u%(k%)) / (u%(k% + T% - 1) - u%(k%)) * SplineBlend(k%, T% - 1, u%(), v&) + _
                (u%(k% + T%) - v&) / (u%(k% + T%) - u%(k% + 1)) * SplineBlend(k% + 1, T% - 1, u%(), v&)
        End If
    End If
    SplineBlend& = value&
End Function

'/*
'   The positions of the subintervals of v and breakpoints, the position
'   on the curve are called knots. Breakpoints can be uniformly defined
'   by setting u[j] = j, a more useful series of breakpoints are defined
'   by the function below. This set of breakpoints localises changes to
'   the vicinity of the control point being modified.
'*/
Sub SplineKnots(u() As Integer, N%, T%)
'int *u,n,t;
'{
'   int j;

'   for (j=0;j<=n+t;j++) {
'      if (j < t)
'         u[j] = 0;
'      else if (j <= n)
'         u[j] = j - t + 1;
'      else if (j > n)
'         u[j] = n - t + 2;
'   }

    Dim j%
    
    For j% = 0 To N% + T% + 1
        If j% < T% Then
            u%(j%) = 0
        ElseIf j% <= N% Then
            u%(j%) = j% - T% + 1
        ElseIf j% > N% Then
            u%(j%) = N% - T% + 2
        End If
    Next j%
End Sub

'/*-------------------------------------------------------------------------
'   Create all the points along a spline curve
'   Control points "inp", "n" of them.
'   Knots "knots", degree "t".
'   Ouput curve "outp", "res" of them.
'*/
Sub SplineCurve(inp() As D3DVECTOR, N%, knots() As Integer, T%, outp() As D3DVECTOR, res%)
'XYZ *inp;
'int n;
'int *knots;
'int t;
'XYZ *outp;
'int res;
'{
'   int i;
'   double interval,increment;

'   interval = 0;
'   increment = (n - t + 2) / (double)(res - 1);
'   for (i=0;i<res-1;i++) {
'      SplinePoint(knots,n,t,interval,inp,&(outp[i]));
'      interval += increment;
'   }
'   outp[res-1] = inp[n];
       
    Dim i%
    Dim interval&, increment&
    
    interval& = 0
    increment& = (N% - T% + 2) / (res% - 1)
    For i% = 0 To res% - 1
        Call SplinePoint(knots%(), N%, T%, interval&, inp(), outp(i))
        interval& = interval& + increment&
    Next i%
    outp(res% - 1) = inp(N%)
End Sub

'
' Example of how to call the spline functions
' Basically one needs to create the control points, then compute
' the knot positions, then calculate points along the curve.
'
Sub main(argc%, argv$)
'XYZ inp[N+1] = {0.0,0.0,0.0,   1.0,0.0,3.0,   2.0,0.0,1.0,   4.0,0.0,4.0};
'int argc;
'char **argv;
'{
'   int i;
'   SplineKnots(knots,N,T);
'   SplineCurve(inp,N,knots,T,outp,RESOLUTION);
'   /* Display the curve, in this case in OOGL format for GeomView */
'   printf("LIST\n");
'   printf("{ = SKEL\n");
'   printf("%d %d\n",RESOLUTION,RESOLUTION-1);
'   for (i=0;i<RESOLUTION;i++)
'      printf("%g %g %g\n",outp[i].x,outp[i].y,outp[i].z);
'   for (i=0;i<RESOLUTION-1;i++)
'      printf("2 %d %d 1 1 1 1\n",i,i+1);
'   printf("}\n");
'   /* The axes */
'   printf("{ = SKEL 3 2  0 0 4  0 0 0  4 0 0  2 0 1 0 0 1 1 2 1 2 0 0 1 1 }\n");
'   /* Control point polygon */
'   printf("{ = SKEL\n");
'   printf("%d %d\n",N+1,N);
'   for (i=0;i<=N;i++)
'      printf("%g %g %g\n",inp[i].x,inp[i].y,inp[i].z);
'   for(i=0;i<N;i++)
'      printf("2 %d %d 0 1 0 1\n",i,i+1);
'   printf("}\n");

    inp(0).X = 0: inp(0).Y = 0: inp(0).z = 0
    inp(1).X = 1: inp(1).Y = 0: inp(1).z = 3
    inp(2).X = 2: inp(2).Y = 0: inp(2).z = 1
    inp(3).X = 4: inp(3).Y = 0: inp(3).z = 4
    
    Dim i%
    Call SplineKnots(knots%(), N%, T%)
For i% = 0 To N% + T% + 1
    Debug.Print knots%(i%)
Next i%
    Call SplineCurve(inp(), N%, knots%(), T%, outp(), RESOLUTION%)
    '   /* Display the curve, in this case in OOGL format for GeomView */
    'Debug.Print "LIST"
    'Debug.Print "{ = SKEL"
    'Debug.Print RESOLUTION%; RESOLUTION% - 1 ' ???
    For i% = 0 To RESOLUTION%
    '    Debug.Print outp(i%).X; outp(i%).Y; outp(i%).z
    Next i%
    For i% = 0 To RESOLUTION% - 1
    '    Debug.Print 2; i%; i% + 1; 1; 1; 1; 1
    Next i%
    'Debug.Print "}"
    '   /* The axes */
    'Debug.Print "{ = SKEL 3 2  0 0 4  0 0 0  4 0 0  2 0 1 0 0 1 1 2 1 2 0 0 1 1 }"
    '   /* Control point polygon */
    'Debug.Print "{ = SKEL"
    'Debug.Print N% + 1; N%
    For i% = 0 To N% + 1
    '    Debug.Print inp(i%).X; inp(i%).Y; inp(i%).z
    Next i%
    For i% = 0 To N%
    '    Debug.Print 2; i%; i% + 1; 0; 1; 0; 1
    Next i%
    'Debug.Print "}"
End Sub

