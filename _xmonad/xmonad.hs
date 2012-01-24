import XMonad
import XMonad.Config.Gnome
import XMonad.ManageHook
import XMonad.Util.EZConfig

myManageHook :: [ManageHook]
myManageHook =
    [ resource  =? "Do"   --> doIgnore ]

main = 
    xmonad $ gnomeConfig
        { 
            manageHook = manageHook gnomeConfig <+> composeAll myManageHook
            , terminal = "gnome-terminal"
            , modMask = mod4Mask
            , focusFollowsMouse = False
            , borderWidth = 2
        }
        `additionalKeysP` 
            [ ("M-S-q", spawn "gnome-session-save --gui --logout-dialog")  -- mod quit logs out of gnome session
            , ("M-S-l",    spawn "gnome-screensaver-command -l")  -- more key bindings to gnome session
            , ("M1-M-S-l", spawn "gnome-session-save --gui --kill")
            , ("M1-S-'",   spawn "gnome-power-cmd.sh suspend")
            , ("M1-S-,",   spawn "gnome-power-cmd.sh reboot")
            , ("M1-S-.",   spawn "gnome-power-cmd.sh hibernate")
            , ("M1-S-p",   spawn "gnome-power-cmd.sh shutdown") 
            ]
