# UserForm:

+ Name: UserForm1
+ Caption: Matrix ops


## Frame:
    
    + Name: frNoOfMatrices
    + Caption: Number of matrices

    OptionButton:
        
        + Name: obSingleMat
        + Caption: Single

    OptionButton:
        
        + Name: obMultipleMat
        + Caption: Multiple



CommandButton:

    + Name: cbSelectCells
    + Caption: Select cells



Label:

    + Name: lblSelectedCells
    + Caption: Cells selected:



Frame:
    
    + Name: frSingle
    + Caption: Single matrix

    CheckBox:

        + Name: cbSingleRank
        + Caption: Rank

    CheckBox:

        + Name: cbSingleLU
        + Caption: LU decomposition

    CheckBox:

        + Name: cbSingleDet
        + Caption: Determinant

    CheckBox:

        + Name: cbSingleNil
        + Caption: Nilpotent

    CheckBox:

        + Name: cbSingleEigen
        + Caption: Eigenvalues & eigenvectors

    CheckBox:

        + Name: cbSinglePow
        + Caption: Power

    Label:
        
        + Name: lblNEquiv
        + Caption: N =

    TextBox:
        
        + Name: tbNPowVal
    
    SpinButton:
        
        + Name: sbNPowVal


    Frame:
        
        + Name: frSingleFunctions
        + Caption: Functions

        Frame:

            + Name: frSingleFuncsTrig
            + Caption: Trigonometric

            CheckBox:
                
                + Name: cbSingleSin
                + Caption: sin

            CheckBox:
                
                + Name: cbSingleCos
                + Caption: cos

            CheckBox:
                
                + Name: cbSingleTan
                + Caption: tan

            CheckBox:
                
                + Name: cbSingleCot
                + Caption: cot



        Frame:

            + Name: frSingleFuncsHyperbolic
            + Caption: Hyperbolic

            CheckBox:
                
                + Name: cbSingleSinh
                + Caption: sinh

            CheckBox:
                
                + Name: cbSingleCosh
                + Caption: cosh

            CheckBox:
                
                + Name: cbSingleTanh
                + Caption: tanh

            CheckBox:
                
                + Name: cbSingleCoth
                + Caption: coth

        Frame:

            + Name: Inverse trigonometric
            + Caption: frSingleFuncsInvTrig

            CheckBox:
                
                + Name: cbSingleArcSin
                + Caption: arcsin

            CheckBox:
                
                + Name: cbSingleArcCos
                + Caption: arccos

            CheckBox:
                
                + Name: cbSingleArcTan
                + Caption: arctan

            CheckBox:
                
                + Name: cbSingleArcCot
                + Caption: arccot

        Frame:

            + Name: Misc.
            + Caption: frSingleFuncsMisc

            CheckBox:
                
                + Name: cbSingleLog
                + Caption: log



Frame:
    
    + Name: frMultiple
    + Caption: Multiple

    CheckBox:
        
        + Name: cbMultipleSum
        + Caption: Sum

    CheckBox:
        
        + Name: cbMultipleProd
        + Caption: Product

    CheckBox:
        
        + Name: cbMultipleDiff
        + Caption: Difference


CommandButton:
    
    + Name: cbSaveChoice
    + Caption: Save choice

CommandButton:
    
    + Name: cbCancel
    + Caption: Cancel

CommandButton:
    
    + Name: cbOK
    + Caption: OK

