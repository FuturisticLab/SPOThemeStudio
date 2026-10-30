<#

.SYNOPSIS

    SharePoint Online Theme Studio & Interactive Visualizer

.DESCRIPTION

    An interactive PowerShell GUI utility for generating, previewing, and deploying custom

    color themes for SharePoint Online (Add-SPOTheme / Set-SPOWebTheme).

    Provides bidirectional inspection: hover or click any element or (i) info icon on the

    SharePoint mockup to see its mapped theme slot, hex code, and role, or modify any slot on the form to see

    the live page update in real-time.

.EXAMPLE

    .\SPOThemeStudio.ps1

.EXAMPLE

    powershell -ExecutionPolicy Bypass -File .\SPOThemeStudio.ps1

#>



[CmdletBinding()]

param()



# Ensure required assemblies are loaded

Add-Type -AssemblyName PresentationFramework

Add-Type -AssemblyName PresentationCore

Add-Type -AssemblyName WindowsBase

Add-Type -AssemblyName System.Drawing

Add-Type -AssemblyName System.Windows.Forms



# ---------------------------------------------------------------------------

# COLOR MATHEMATICS & UTILITY FUNCTIONS

# ---------------------------------------------------------------------------



function Convert-HexToDrawingColor([string]$hex) {

    try {

        $clean = $hex.Trim().TrimStart('#')

        if ($clean.Length -eq 3) {

            $clean = "$($clean[0])$($clean[0])$($clean[1])$($clean[1])$($clean[2])$($clean[2])"

        }

        if ($clean.Length -eq 6) {

            $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)

            $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)

            $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)

            return [System.Drawing.Color]::FromArgb(255, $r, $g, $b)

        }

    } catch {}

    return [System.Drawing.Color]::FromArgb(255, 0, 120, 212)

}



function Convert-DrawingColorToHex([System.Drawing.Color]$c) {

    return ('#{0:x2}{1:x2}{2:x2}' -f $c.R, $c.G, $c.B).ToLower()

}



function Convert-HexToMediaBrush([string]$hex) {

    try {

        $c = Convert-HexToDrawingColor $hex

        $mediaColor = [System.Windows.Media.Color]::FromArgb($c.A, $c.R, $c.G, $c.B)

        $brush = [System.Windows.Media.SolidColorBrush]::new($mediaColor)

        $brush.Freeze()

        return $brush

    } catch {

        return [System.Windows.Media.Brushes]::Transparent

    }

}



function Blend-HexColors([string]$hex1, [string]$hex2, [double]$ratio) {

    $c1 = Convert-HexToDrawingColor $hex1

    $c2 = Convert-HexToDrawingColor $hex2

    $r = [Math]::Max(0, [Math]::Min(255, [int][Math]::Round($c1.R * (1.0 - $ratio) + $c2.R * $ratio)))

    $g = [Math]::Max(0, [Math]::Min(255, [int][Math]::Round($c1.G * (1.0 - $ratio) + $c2.G * $ratio)))

    $b = [Math]::Max(0, [Math]::Min(255, [int][Math]::Round($c1.B * (1.0 - $ratio) + $c2.B * $ratio)))

    return ('#{0:x2}{1:x2}{2:x2}' -f $r, $g, $b).ToLower()

}



function Convert-RgbToHsl($r, $g, $b) {

    $rNorm = $r / 255.0

    $gNorm = $g / 255.0

    $bNorm = $b / 255.0



    $max = [Math]::Max($rNorm, [Math]::Max($gNorm, $bNorm))

    $min = [Math]::Min($rNorm, [Math]::Min($gNorm, $bNorm))

    $delta = $max - $min



    $l = ($max + $min) / 2.0

    $h = 0.0

    $s = 0.0



    if ($delta -gt 0.00001) {

        if ($l -lt 0.5) {

            $s = $delta / ($max + $min)

        } else {

            $s = $delta / (2.0 - $max - $min)

        }



        if ($max -eq $rNorm) {

            $h = (($gNorm - $bNorm) / $delta)

            if ($gNorm -lt $bNorm) { $h += 6.0 }

        } elseif ($max -eq $gNorm) {

            $h = (($bNorm - $rNorm) / $delta) + 2.0

        } else {

            $h = (($rNorm - $gNorm) / $delta) + 4.0

        }

        $h *= 60.0

    }



    return @{ H = $h; S = $s; L = $l }

}



function HueToRgb($p, $q, $t) {

    if ($t -lt 0.0) { $t += 1.0 }

    if ($t -gt 1.0) { $t -= 1.0 }

    if ($t -lt (1.0 / 6.0)) { return $p + ($q - $p) * 6.0 * $t }

    if ($t -lt 0.5) { return $q }

    if ($t -lt (2.0 / 3.0)) { return $p + ($q - $p) * (2.0 / 3.0 - $t) * 6.0 }

    return $p

}



function Convert-HslToHex($h, $s, $l) {

    $hNorm = ($h % 360.0) / 360.0

    if ($s -eq 0.0) {

        $gray = [int][Math]::Round($l * 255.0)

        return ('#{0:x2}{1:x2}{2:x2}' -f $gray, $gray, $gray)

    }



    $q = if ($l -lt 0.5) { $l * (1.0 + $s) } else { $l + $s - ($l * $s) }

    $p = 2.0 * $l - $q



    $r = [int][Math]::Round((HueToRgb $p $q ($hNorm + (1.0 / 3.0))) * 255.0)

    $g = [int][Math]::Round((HueToRgb $p $q $hNorm) * 255.0)

    $b = [int][Math]::Round((HueToRgb $p $q ($hNorm - (1.0 / 3.0))) * 255.0)



    $r = [Math]::Max(0, [Math]::Min(255, $r))

    $g = [Math]::Max(0, [Math]::Min(255, $g))

    $b = [Math]::Max(0, [Math]::Min(255, $b))



    return ('#{0:x2}{1:x2}{2:x2}' -f $r, $g, $b).ToLower()

}



# ---------------------------------------------------------------------------

# SHAREPOINT THEME SLOTS DEFINITION & METADATA

# ---------------------------------------------------------------------------



$script:SlotMetadata = [ordered]@{

    # --- Theme Primary / Brand Accents ---

    "themePrimary" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Primary"

        Role = "Main brand accent. Used for primary action buttons ('+ New', 'Submit'), active tab underlines, hyperlinked text, checkboxes, suite bar accents, and hero banner CTA buttons."

        UIElements = @("suiteBarAccent", "btnPrimary", "heroBtnPrimary", "navActiveLine", "linkSample", "quickLinkIcon1", "quickLinkIcon2", "quickLinkIcon3")

    }

    "themeLighterAlt" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Lighter Alt"

        Role = "Extremely light tint. Used for subtle hover backgrounds on list items, selected row background highlights, and command bar active states."

        UIElements = @("tableRowSelected")

    }

    "themeLighter" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Lighter"

        Role = "Light tint. Used for hover highlights on secondary pills, progress bar background tracks, and soft tag fills."

        UIElements = @("pillHighlight")

    }

    "themeLight" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Light"

        Role = "Medium-light tint. Used for command bar button hover fills, subtle alert banners, and active filter chip backgrounds."

        UIElements = @("commandBarBtnHover")

    }

    "themeTertiary" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Tertiary"

        Role = "Mid-tone tint. Used for hyperlink hover color, secondary progress bars, tag outlines, and slider active tracks."

        UIElements = @("linkHover", "heroTagChip")

    }

    "themeSecondary" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Secondary"

        Role = "Slightly lighter primary shade. Used for secondary button outlines, primary button hover states, and active folder icons."

        UIElements = @("folderIcon1", "folderIcon2")

    }

    "themeDarkAlt" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Dark Alt"

        Role = "Dark primary shade. Used for primary button hover background, hero banner gradient shading, and active dropdown items."

        UIElements = @("heroBannerBg")

    }

    "themeDark" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Dark"

        Role = "Dark primary shade. Used for primary button pressed state, focused text input borders, and active navigation tab text."

        UIElements = @("navActiveText")

    }

    "themeDarker" = @{

        Category = "Brand Accent"

        DisplayName = "Theme Darker"

        Role = "Deepest primary shade. Used for pressed link states and maximum-contrast brand text on light surfaces."

        UIElements = @("linkPressed")

    }



    # --- Neutrals & Canvas ---

    "neutralLighterAlt" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Lighter Alt"

        Role = "Subtle background surface. Used for card container surfaces, web part canvas containers, and alternating table rows."

        UIElements = @("quickLinkCard1", "quickLinkCard2", "quickLinkCard3", "docLibraryCard")

    }

    "neutralLighter" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Lighter"

        Role = "Default page canvas background in standard light themes. Secondary button backgrounds and command bar fill."

        UIElements = @("pageBodyCanvas", "btnSecondary", "heroBtnSecondary")

    }

    "neutralLight" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Light"

        Role = "Divider lines, subtle card borders, table column separators, and card hover fills."

        UIElements = @("dividerLine1", "dividerLine2", "cardBorder1", "cardBorder2", "cardBorder3")

    }

    "neutralQuaternaryAlt" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Quaternary Alt"

        Role = "Table borders, list separator lines, and secondary disabled checkbox borders."

        UIElements = @("tableRowBorder")

    }

    "neutralQuaternary" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Quaternary"

        Role = "Input borders (search box, text inputs), secondary card borders, and inactive slider tracks."

        UIElements = @("searchBoxBorder", "btnSecondaryBorder")

    }

    "neutralTertiaryAlt" = @{

        Category = "Neutral Surface"

        DisplayName = "Neutral Tertiary Alt"

        Role = "Disabled icons, placeholder glyphs, muted tag backgrounds, and inactive indicators."

        UIElements = @("disabledIcon", "disabledTagBg")

    }

    "neutralTertiary" = @{

        Category = "Neutral Text"

        DisplayName = "Neutral Tertiary"

        Role = "Disabled text on buttons, search box placeholder text, and muted timestamp metadata."

        UIElements = @("searchPlaceholder", "disabledBtnText")

    }

    "neutralSecondary" = @{

        Category = "Neutral Text"

        DisplayName = "Neutral Secondary"

        Role = "Secondary metadata text ('Modified 2 hours ago by admin'), subtitles, breadcrumbs, and inactive nav tabs."

        UIElements = @("metaTimestamp", "navInactiveText", "tableHeaderLabel", "cardDesc1", "cardDesc2", "cardDesc3")

    }

    "neutralPrimaryAlt" = @{

        Category = "Neutral Text"

        DisplayName = "Neutral Primary Alt"

        Role = "Hover text for dark titles, sub-headings, and strong metadata labels."

        UIElements = @("newsSubHeading")

    }

    "neutralPrimary" = @{

        Category = "Neutral Text"

        DisplayName = "Neutral Primary"

        Role = "Main body paragraph text, primary list labels, document titles, and web part headings."

        UIElements = @("siteTitleText", "bodyParagraphText", "docName1", "docName2", "cardTitle1", "cardTitle2", "cardTitle3", "btnSecondaryText")

    }

    "neutralDark" = @{

        Category = "Neutral Text"

        DisplayName = "Neutral Dark"

        Role = "Maximum contrast headings, hero title text, page H1 headings, and strong bold titles."

        UIElements = @("pageMainHeading", "webPartSectionTitle")

    }

    "black" = @{

        Category = "Neutral Core"

        DisplayName = "Black"

        Role = "Deep black text and high-contrast elements. In inverted dark themes, mapped to high-contrast white."

        UIElements = @("blackHighContrast")

    }

    "white" = @{

        Category = "Neutral Core"

        DisplayName = "White"

        Role = "Standard page canvas surface, card surfaces, and primary button text. In inverted themes, mapped to deep dark canvas."

        UIElements = @("suiteBarHeader", "btnPrimaryText", "heroBtnPrimaryText", "siteHeaderSurface")

    }



    # --- Extended Semantic Slots ---

    "primaryBackground" = @{

        Category = "Semantic Extension"

        DisplayName = "Primary Background"

        Role = "Primary application background surface. Equivalent to white in light theme or dark onyx in inverted theme."

        UIElements = @("pageBodyCanvas")

    }

    "primaryText" = @{

        Category = "Semantic Extension"

        DisplayName = "Primary Text"

        Role = "Primary application text color. Equivalent to neutralPrimary."

        UIElements = @("bodyParagraphText")

    }

    "bodyBackground" = @{

        Category = "Semantic Extension"

        DisplayName = "Body Background"

        Role = "Container body surface background."

        UIElements = @("pageBodyCanvas")

    }

    "bodyText" = @{

        Category = "Semantic Extension"

        DisplayName = "Body Text"

        Role = "Paragraph and body font text color."

        UIElements = @("bodyParagraphText")

    }

    "disabledBackground" = @{

        Category = "Semantic Extension"

        DisplayName = "Disabled Background"

        Role = "Background color for inactive and disabled buttons and form controls."

        UIElements = @("disabledTagBg")

    }

    "disabledText" = @{

        Category = "Semantic Extension"

        DisplayName = "Disabled Text"

        Role = "Foreground label color for disabled controls."

        UIElements = @("disabledBtnText")

    }

}



# ---------------------------------------------------------------------------

# PRESET PALETTES

# ---------------------------------------------------------------------------



$script:Presets = @{

    "Microsoft Modern Blue" = @{

        IsInverted = $false

        Slots = @{

            "themePrimary"         = "#0078d4"

            "themeLighterAlt"      = "#eff6fc"

            "themeLighter"         = "#deecf9"

            "themeLight"           = "#c7e0f4"

            "themeTertiary"        = "#71afe5"

            "themeSecondary"       = "#2b88d8"

            "themeDarkAlt"         = "#106ebe"

            "themeDark"            = "#005a9e"

            "themeDarker"          = "#004578"

            "neutralLighterAlt"    = "#faf9f8"

            "neutralLighter"       = "#f3f2f1"

            "neutralLight"         = "#edebe9"

            "neutralQuaternaryAlt" = "#e1dfdd"

            "neutralQuaternary"    = "#d2d0ce"

            "neutralTertiaryAlt"   = "#c8c6c4"

            "neutralTertiary"      = "#a19f9d"

            "neutralSecondary"     = "#605e5c"

            "neutralPrimaryAlt"    = "#3b3a39"

            "neutralPrimary"       = "#323130"

            "neutralDark"          = "#201f1e"

            "black"                = "#000000"

            "white"                = "#ffffff"

            "primaryBackground"    = "#ffffff"

            "primaryText"          = "#323130"

            "bodyBackground"       = "#ffffff"

            "bodyText"             = "#323130"

            "disabledBackground"   = "#f3f2f1"

            "disabledText"         = "#a19f9d"

        }

    }

    "Cyberdyne Crimson HUD" = @{

        IsInverted = $true

        Slots = @{

            "themePrimary"         = "#ff1e3c"

            "themeLighterAlt"      = "#0a0103"

            "themeLighter"         = "#29050a"

            "themeLight"           = "#4d0912"

            "themeTertiary"        = "#991224"

            "themeSecondary"       = "#e01a35"

            "themeDarkAlt"         = "#ff3550"

            "themeDark"            = "#ff546b"

            "themeDarker"          = "#ff8293"

            "neutralLighterAlt"    = "#111319"

            "neutralLighter"       = "#161922"

            "neutralLight"         = "#1f2430"

            "neutralQuaternaryAlt" = "#262b38"

            "neutralQuaternary"    = "#2c3242"

            "neutralTertiaryAlt"   = "#444c63"

            "neutralTertiary"      = "#c8c8c8"

            "neutralSecondary"     = "#d0d0d0"

            "neutralPrimaryAlt"    = "#dadada"

            "neutralPrimary"       = "#ffffff"

            "neutralDark"          = "#f4f4f4"

            "black"                = "#f8f8f8"

            "white"                = "#090a0d"

            "primaryBackground"    = "#090a0d"

            "primaryText"          = "#ffffff"

            "bodyBackground"       = "#090a0d"

            "bodyText"             = "#ffffff"

            "disabledBackground"   = "#161922"

            "disabledText"         = "#636b7b"

        }

    }

    "Emerald Forest" = @{

        IsInverted = $false

        Slots = @{

            "themePrimary"         = "#107c41"

            "themeLighterAlt"      = "#f1faf3"

            "themeLighter"         = "#c8ebd3"

            "themeLight"           = "#9ddab0"

            "themeTertiary"        = "#51b573"

            "themeSecondary"       = "#218c4e"

            "themeDarkAlt"         = "#0e703a"

            "themeDark"            = "#0c5e31"

            "themeDarker"          = "#094624"

            "neutralLighterAlt"    = "#faf9f8"

            "neutralLighter"       = "#f3f2f1"

            "neutralLight"         = "#edebe9"

            "neutralQuaternaryAlt" = "#e1dfdd"

            "neutralQuaternary"    = "#d2d0ce"

            "neutralTertiaryAlt"   = "#c8c6c4"

            "neutralTertiary"      = "#a19f9d"

            "neutralSecondary"     = "#605e5c"

            "neutralPrimaryAlt"    = "#3b3a39"

            "neutralPrimary"       = "#323130"

            "neutralDark"          = "#201f1e"

            "black"                = "#000000"

            "white"                = "#ffffff"

            "primaryBackground"    = "#ffffff"

            "primaryText"          = "#323130"

            "bodyBackground"       = "#ffffff"

            "bodyText"             = "#323130"

            "disabledBackground"   = "#f3f2f1"

            "disabledText"         = "#a19f9d"

        }

    }

    "Teal Horizon" = @{

        IsInverted = $false

        Slots = @{

            "themePrimary"         = "#038387"

            "themeLighterAlt"      = "#f1fbfc"

            "themeLighter"         = "#caf0f1"

            "themeLight"           = "#9ee2e4"

            "themeTertiary"        = "#4dc2c6"

            "themeSecondary"       = "#16979b"

            "themeDarkAlt"         = "#03767a"

            "themeDark"            = "#026467"

            "themeDarker"          = "#024a4c"

            "neutralLighterAlt"    = "#faf9f8"

            "neutralLighter"       = "#f3f2f1"

            "neutralLight"         = "#edebe9"

            "neutralQuaternaryAlt" = "#e1dfdd"

            "neutralQuaternary"    = "#d2d0ce"

            "neutralTertiaryAlt"   = "#c8c6c4"

            "neutralTertiary"      = "#a19f9d"

            "neutralSecondary"     = "#605e5c"

            "neutralPrimaryAlt"    = "#3b3a39"

            "neutralPrimary"       = "#323130"

            "neutralDark"          = "#201f1e"

            "black"                = "#000000"

            "white"                = "#ffffff"

            "primaryBackground"    = "#ffffff"

            "primaryText"          = "#323130"

            "bodyBackground"       = "#ffffff"

            "bodyText"             = "#323130"

            "disabledBackground"   = "#f3f2f1"

            "disabledText"         = "#a19f9d"

        }

    }

    "Midnight Amethyst" = @{

        IsInverted = $false

        Slots = @{

            "themePrimary"         = "#7030a0"

            "themeLighterAlt"      = "#f8f4fb"

            "themeLighter"         = "#e4d6f0"

            "themeLight"           = "#cdb4e3"

            "themeTertiary"        = "#a074c7"

            "themeSecondary"       = "#7c41ab"

            "themeDarkAlt"         = "#652b90"

            "themeDark"            = "#55247a"

            "themeDarker"          = "#3f1b5a"

            "neutralLighterAlt"    = "#faf9f8"

            "neutralLighter"       = "#f3f2f1"

            "neutralLight"         = "#edebe9"

            "neutralQuaternaryAlt" = "#e1dfdd"

            "neutralQuaternary"    = "#d2d0ce"

            "neutralTertiaryAlt"   = "#c8c6c4"

            "neutralTertiary"      = "#a19f9d"

            "neutralSecondary"     = "#605e5c"

            "neutralPrimaryAlt"    = "#3b3a39"

            "neutralPrimary"       = "#323130"

            "neutralDark"          = "#201f1e"

            "black"                = "#000000"

            "white"                = "#ffffff"

            "primaryBackground"    = "#ffffff"

            "primaryText"          = "#323130"

            "bodyBackground"       = "#ffffff"

            "bodyText"             = "#323130"

            "disabledBackground"   = "#f3f2f1"

            "disabledText"         = "#a19f9d"

        }

    }

    "Sunset Amber" = @{

        IsInverted = $false

        Slots = @{

            "themePrimary"         = "#d83b01"

            "themeLighterAlt"      = "#fdf5f2"

            "themeLighter"         = "#f8d8cc"

            "themeLight"           = "#f2b7a2"

            "themeTertiary"        = "#e5734e"

            "themeSecondary"       = "#dc4c16"

            "themeDarkAlt"         = "#c33401"

            "themeDark"            = "#a52d01"

            "themeDarker"          = "#7a2101"

            "neutralLighterAlt"    = "#faf9f8"

            "neutralLighter"       = "#f3f2f1"

            "neutralLight"         = "#edebe9"

            "neutralQuaternaryAlt" = "#e1dfdd"

            "neutralQuaternary"    = "#d2d0ce"

            "neutralTertiaryAlt"   = "#c8c6c4"

            "neutralTertiary"      = "#a19f9d"

            "neutralSecondary"     = "#605e5c"

            "neutralPrimaryAlt"    = "#3b3a39"

            "neutralPrimary"       = "#323130"

            "neutralDark"          = "#201f1e"

            "black"                = "#000000"

            "white"                = "#ffffff"

            "primaryBackground"    = "#ffffff"

            "primaryText"          = "#323130"

            "bodyBackground"       = "#ffffff"

            "bodyText"             = "#323130"

            "disabledBackground"   = "#f3f2f1"

            "disabledText"         = "#a19f9d"

        }

    }

    "Obsidian Dark Executive" = @{

        IsInverted = $true

        Slots = @{

            "themePrimary"         = "#00b4d8"

            "themeLighterAlt"      = "#001014"

            "themeLighter"         = "#002830"

            "themeLight"           = "#004a5c"

            "themeTertiary"        = "#007f9c"

            "themeSecondary"       = "#00a1c2"

            "themeDarkAlt"         = "#1fc1e0"

            "themeDark"            = "#4bd0e8"

            "themeDarker"          = "#82e2f2"

            "neutralLighterAlt"    = "#14171f"

            "neutralLighter"       = "#1b202b"

            "neutralLight"         = "#242a38"

            "neutralQuaternaryAlt" = "#2c3344"

            "neutralQuaternary"    = "#343d52"

            "neutralTertiaryAlt"   = "#4f5a75"

            "neutralTertiary"      = "#c4c8d4"

            "neutralSecondary"     = "#d2d6e0"

            "neutralPrimaryAlt"    = "#e2e5ec"

            "neutralPrimary"       = "#f8f9fb"

            "neutralDark"          = "#ffffff"

            "black"                = "#ffffff"

            "white"                = "#0c0e14"

            "primaryBackground"    = "#0c0e14"

            "primaryText"          = "#f8f9fb"

            "bodyBackground"       = "#0c0e14"

            "bodyText"             = "#f8f9fb"

            "disabledBackground"   = "#1b202b"

            "disabledText"         = "#5a6582"

        }

    }

}



# ---------------------------------------------------------------------------

# PALETTE CALCULATION & SCHEME GENERATION ALGORITHMS

# ---------------------------------------------------------------------------



function Generate-PaletteFromPrimary([string]$primaryHex, [bool]$inverted = $false) {

    $primaryHex = $primaryHex.ToLower()

    $palette = [ordered]@{}

    $palette["themePrimary"] = $primaryHex



    if (-not $inverted) {

        # Light theme shade calculation

        $palette["themeLighterAlt"] = Blend-HexColors $primaryHex "#ffffff" 0.94

        $palette["themeLighter"]    = Blend-HexColors $primaryHex "#ffffff" 0.82

        $palette["themeLight"]      = Blend-HexColors $primaryHex "#ffffff" 0.65

        $palette["themeTertiary"]   = Blend-HexColors $primaryHex "#ffffff" 0.45

        $palette["themeSecondary"]  = Blend-HexColors $primaryHex "#ffffff" 0.18

        $palette["themeDarkAlt"]    = Blend-HexColors $primaryHex "#000000" 0.12

        $palette["themeDark"]       = Blend-HexColors $primaryHex "#000000" 0.25

        $palette["themeDarker"]     = Blend-HexColors $primaryHex "#000000" 0.42



        $palette["neutralLighterAlt"] = "#faf9f8"

        $palette["neutralLighter"]    = "#f3f2f1"

        $palette["neutralLight"]      = "#edebe9"

        $palette["neutralQuaternaryAlt"] = "#e1dfdd"

        $palette["neutralQuaternary"]    = "#d2d0ce"

        $palette["neutralTertiaryAlt"]   = "#c8c6c4"

        $palette["neutralTertiary"]      = "#a19f9d"

        $palette["neutralSecondary"]     = "#605e5c"

        $palette["neutralPrimaryAlt"]    = "#3b3a39"

        $palette["neutralPrimary"]       = "#323130"

        $palette["neutralDark"]          = "#201f1e"

        $palette["black"]                = "#000000"

        $palette["white"]                = "#ffffff"



        $palette["primaryBackground"]  = "#ffffff"

        $palette["primaryText"]        = "#323130"

        $palette["bodyBackground"]     = "#ffffff"

        $palette["bodyText"]           = "#323130"

        $palette["disabledBackground"] = "#f3f2f1"

        $palette["disabledText"]       = "#a19f9d"

    } else {

        # Inverted / Dark theme calculation

        $palette["themeLighterAlt"] = Blend-HexColors $primaryHex "#000000" 0.92

        $palette["themeLighter"]    = Blend-HexColors $primaryHex "#000000" 0.78

        $palette["themeLight"]      = Blend-HexColors $primaryHex "#000000" 0.60

        $palette["themeTertiary"]   = Blend-HexColors $primaryHex "#000000" 0.40

        $palette["themeSecondary"]  = Blend-HexColors $primaryHex "#000000" 0.15

        $palette["themeDarkAlt"]    = Blend-HexColors $primaryHex "#ffffff" 0.15

        $palette["themeDark"]       = Blend-HexColors $primaryHex "#ffffff" 0.30

        $palette["themeDarker"]     = Blend-HexColors $primaryHex "#ffffff" 0.50



        $palette["neutralLighterAlt"] = "#111319"

        $palette["neutralLighter"]    = "#161922"

        $palette["neutralLight"]      = "#1f2430"

        $palette["neutralQuaternaryAlt"] = "#262b38"

        $palette["neutralQuaternary"]    = "#2c3242"

        $palette["neutralTertiaryAlt"]   = "#444c63"

        $palette["neutralTertiary"]      = "#c8c8c8"

        $palette["neutralSecondary"]     = "#d0d0d0"

        $palette["neutralPrimaryAlt"]    = "#dadada"

        $palette["neutralPrimary"]       = "#ffffff"

        $palette["neutralDark"]          = "#f4f4f4"

        $palette["black"]                = "#f8f8f8"

        $palette["white"]                = "#090a0d"



        $palette["primaryBackground"]  = "#090a0d"

        $palette["primaryText"]        = "#ffffff"

        $palette["bodyBackground"]     = "#090a0d"

        $palette["bodyText"]           = "#ffffff"

        $palette["disabledBackground"] = "#161922"

        $palette["disabledText"]       = "#636b7b"

    }



    return $palette

}



function Generate-RandomScheme([bool]$inverted = $false) {

    $rand = [System.Random]::new()

    $hue = $rand.Next(0, 360)

    $sat = ($rand.Next(65, 95)) / 100.0

    $light = if ($inverted) { ($rand.Next(45, 60)) / 100.0 } else { ($rand.Next(38, 52)) / 100.0 }



    $primaryHex = Convert-HslToHex $hue $sat $light

    return Generate-PaletteFromPrimary $primaryHex $inverted

}



# ---------------------------------------------------------------------------

# WPF USER INTERFACE DEFINITION (XAML)

# ---------------------------------------------------------------------------



$xamlString = @"

<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

        Title="SharePoint Online Theme Studio and Interactive Visualizer"

        Height="920" Width="1460" MinHeight="740" MinWidth="1200"

        WindowStartupLocation="CenterScreen"

        Background="#181820" Foreground="#E6E8EE" FontFamily="Segoe UI">



    <Window.Resources>

        <Style TargetType="TextBlock">

            <Setter Property="Foreground" Value="#E0E3EA"/>

        </Style>

        <Style TargetType="TextBox">

            <Setter Property="Background" Value="#252530"/>

            <Setter Property="Foreground" Value="#FFFFFF"/>

            <Setter Property="BorderBrush" Value="#3D3E4F"/>

            <Setter Property="BorderThickness" Value="1"/>

            <Setter Property="Padding" Value="6,4"/>

            <Setter Property="VerticalContentAlignment" Value="Center"/>

        </Style>

        <Style TargetType="Button">

            <Setter Property="Background" Value="#2E3040"/>

            <Setter Property="Foreground" Value="#FFFFFF"/>

            <Setter Property="BorderBrush" Value="#45475E"/>

            <Setter Property="BorderThickness" Value="1"/>

            <Setter Property="Padding" Value="10,6"/>

            <Setter Property="Cursor" Value="Hand"/>

        </Style>

        <Style TargetType="CheckBox">

            <Setter Property="Foreground" Value="#E0E3EA"/>

            <Setter Property="VerticalContentAlignment" Value="Center"/>

            <Setter Property="Cursor" Value="Hand"/>

        </Style>

            <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="BorderBrush" Value="#60647D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#E2E5EC"/>
                    <Setter Property="Foreground" Value="#000000"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#0078D4"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>



    <Grid>

        <Grid.RowDefinitions>

            <RowDefinition Height="Auto"/> <!-- Top App Header -->

            <RowDefinition Height="*"/>    <!-- Main Body Work Area -->

            <RowDefinition Height="Auto"/> <!-- Status Bar -->

        </Grid.RowDefinitions>



        <!-- TOP APP HEADER BAR -->

        <Border Grid.Row="0" Background="#1F202B" BorderBrush="#2D2E3D" BorderThickness="0,0,0,1" Padding="18,12">

            <Grid>

                <Grid.ColumnDefinitions>

                    <ColumnDefinition Width="Auto"/>

                    <ColumnDefinition Width="*"/>

                    <ColumnDefinition Width="Auto"/>

                </Grid.ColumnDefinitions>



                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">

                    <Border Width="34" Height="34" CornerRadius="6" Background="#0078D4" Margin="0,0,12,0">

                        <TextBlock Text="SP" FontSize="14" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                    </Border>

                    <StackPanel>

                        <TextBlock Text="SharePoint Online Theme Studio" FontSize="17" FontWeight="Bold" Foreground="#FFFFFF"/>

                        <TextBlock Text="Interactive Fluent UI Palette Builder and Component Map" FontSize="11" Foreground="#9BA1B5"/>

                    </StackPanel>

                </StackPanel>



                <!-- TOP QUICK ACTIONS / SCHEME GENERATORS -->

                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">

                    <TextBlock Text="Presets:" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#B0B5C6" FontSize="12"/>

                    <ComboBox x:Name="ComboPresets" Width="195" Height="30" Margin="0,0,10,0" Background="#FFFFFF" Foreground="#000000" BorderBrush="#60647D" FontWeight="SemiBold"/>



                    <Button x:Name="BtnRandomize" Margin="0,0,10,0" Height="30" Background="#3B2E58" BorderBrush="#684E9C" ToolTip="Generate a fresh harmonious palette combination">

                        <TextBlock Text="Roll Random Scheme" FontWeight="SemiBold"/>

                    </Button>



                    <Button x:Name="BtnInvertMode" Margin="0,0,14,0" Height="30" Background="#2A3347" BorderBrush="#425578" ToolTip="Toggle Inverted / Dark Theme">

                        <TextBlock x:Name="TxtInvertMode" Text="Toggle Dark Mode" FontWeight="SemiBold"/>

                    </Button>



                    <Border Width="1" Background="#36384C" Margin="0,2,14,2"/>



                    <Button x:Name="BtnFetchTenant" Margin="0,0,10,0" Height="30" Background="#1D3C58" BorderBrush="#2D608C" ToolTip="Connect to SharePoint and fetch the current theme palette from your tenant">

                        <TextBlock Text="Fetch from Tenant" FontWeight="SemiBold"/>

                    </Button>



                    <Button x:Name="BtnTopCopyScript" Height="30" Background="#0E6038" BorderBrush="#1AA35E" ToolTip="Copy complete Add-SPOTheme PowerShell script">

                        <TextBlock Text="Copy Add-SPOTheme Script" FontWeight="Bold"/>

                    </Button>

                </StackPanel>

            </Grid>

        </Border>



        <!-- MAIN BODY (LEFT FORM / RIGHT INTERACTIVE PREVIEW) -->

        <Grid Grid.Row="1">

            <Grid.ColumnDefinitions>

                <ColumnDefinition Width="480" MinWidth="430" MaxWidth="620"/>

                <ColumnDefinition Width="Auto"/>

                <ColumnDefinition Width="*"/>

            </Grid.ColumnDefinitions>



            <!-- LEFT PANE: FORM & PALETTE FIELDS -->

            <Grid Grid.Column="0" Background="#1A1B24">

                <Grid.RowDefinitions>

                    <RowDefinition Height="Auto"/> <!-- Deployment Metadata -->

                    <RowDefinition Height="Auto"/> <!-- App Registration / Tenant ID (Optional) -->

                    <RowDefinition Height="Auto"/> <!-- Base Color Fast Generator -->

                    <RowDefinition Height="Auto"/> <!-- Search & Filter -->

                    <RowDefinition Height="*"/>    <!-- Scrollable Slot Fields -->

                    <RowDefinition Height="Auto"/> <!-- Action Footer Buttons -->

                </Grid.RowDefinitions>



                <!-- 1. DEPLOYMENT METADATA -->

                <Border Grid.Row="0" Background="#20212C" BorderBrush="#2D2F3F" BorderThickness="0,0,0,1" Padding="14,10">

                    <StackPanel>

                        <TextBlock Text="SHAREPOINT THEME METADATA" FontSize="11" FontWeight="Bold" Foreground="#858BA2" Margin="0,0,0,6"/>

                        <Grid Margin="0,0,0,5">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Theme Identity:" VerticalAlignment="Center" Foreground="#C8CBD8" FontSize="12"/>

                            <TextBox x:Name="TxtThemeName" Grid.Column="1" Text="Contoso Enterprise Blue"/>

                        </Grid>



                        <Grid Margin="0,0,0,5">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Admin URL:" VerticalAlignment="Center" Foreground="#C8CBD8" FontSize="12"/>

                            <TextBox x:Name="TxtAdminUrl" Grid.Column="1" Text="https://YOURTENANT-admin.sharepoint.com"/>

                        </Grid>



                        <Grid Margin="0,0,0,6">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Target Site URL:" VerticalAlignment="Center" Foreground="#C8CBD8" FontSize="12"/>

                            <TextBox x:Name="TxtSiteUrl" Grid.Column="1" Text="https://YOURTENANT.sharepoint.com/sites/portal"/>

                        </Grid>



                        <StackPanel Orientation="Horizontal">

                            <CheckBox x:Name="ChkIsInverted" Content="IsInverted ($true / $false)" Margin="0,0,16,0" IsChecked="False"/>

                            <CheckBox x:Name="ChkOverwrite" Content="Safe Auto-Update (-Overwrite)" IsChecked="True"/>

                        </StackPanel>

                    </StackPanel>

                </Border>



                <!-- 2. OPTIONAL APP REGISTRATION / TENANT CREDENTIALS -->

                <Expander Grid.Row="1" Background="#1D1E2A" BorderBrush="#2E3042" BorderThickness="0,0,0,1" Padding="10,4" IsExpanded="False">

                    <Expander.Header>

                        <TextBlock Text="App-Only / Service Principal Credentials (Optional)" FontSize="11" FontWeight="SemiBold" Foreground="#9BA2BC"/>

                    </Expander.Header>

                    <StackPanel Margin="0,6,0,4">

                        <TextBlock Text="Fill these if using automated Entra ID App registration or unattended CI/CD pipelines:" FontSize="10" Foreground="#7E849B" Margin="0,0,0,6" TextWrapping="Wrap"/>

                        

                        <Grid Margin="0,0,0,5">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Tenant ID:" VerticalAlignment="Center" Foreground="#A9AFCE" FontSize="11"/>

                            <TextBox x:Name="TxtTenantId" Grid.Column="1" FontSize="11" Height="24"/>

                        </Grid>



                        <Grid Margin="0,0,0,5">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Client / App ID:" VerticalAlignment="Center" Foreground="#A9AFCE" FontSize="11"/>

                            <TextBox x:Name="TxtClientId" Grid.Column="1" FontSize="11" Height="24"/>

                        </Grid>



                        <Grid Margin="0,0,0,4">

                            <Grid.ColumnDefinitions>

                                <ColumnDefinition Width="110"/>

                                <ColumnDefinition Width="*"/>

                            </Grid.ColumnDefinitions>

                            <TextBlock Text="Client Secret / Key:" VerticalAlignment="Center" Foreground="#A9AFCE" FontSize="11"/>

                            <PasswordBox x:Name="TxtClientSecret" Grid.Column="1" FontSize="11" Height="24"/>

                        </Grid>

                    </StackPanel>

                </Expander>



                <!-- 3. BASE COLOR QUICK GENERATOR -->

                <Border Grid.Row="2" Background="#242633" BorderBrush="#323447" BorderThickness="0,0,0,1" Padding="14,8">

                    <Grid>

                        <Grid.ColumnDefinitions>

                            <ColumnDefinition Width="Auto"/>

                            <ColumnDefinition Width="*"/>

                            <ColumnDefinition Width="Auto"/>

                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Column="0" Text="Base Primary Hex:" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#D0D4E4" FontSize="12" FontWeight="SemiBold"/>

                        <TextBox x:Name="TxtQuickBaseHex" Grid.Column="1" Text="#0078d4" Margin="0,0,8,0" MaxLength="7" FontFamily="Consolas"/>

                        <Button x:Name="BtnGenerateFromBase" Grid.Column="2" Background="#0F548C" BorderBrush="#1C78C2" Padding="10,4" ToolTip="Auto-calculate all 22 tints and shades from this primary color">

                            <TextBlock Text="Calculate Palette" FontWeight="SemiBold" FontSize="11"/>

                        </Button>

                    </Grid>

                </Border>



                <!-- 4. SEARCH & FILTER BAR -->

                <Border Grid.Row="3" Background="#1E1F2A" Padding="14,6" BorderBrush="#2D2F3F" BorderThickness="0,0,0,1">

                    <Grid>

                        <TextBox x:Name="TxtFilterSlots" Height="26" Padding="6,2,6,2" VerticalContentAlignment="Center" Background="#292A38" BorderBrush="#3C3E52"/>

                        <TextBlock x:Name="TxtFilterPlaceholder" Text="Filter slots (e.g. primary, neutral, border, button)..." Margin="8,0,0,0" VerticalAlignment="Center" Foreground="#7A8096" IsHitTestVisible="False"/>

                    </Grid>

                </Border>



                <!-- 5. SCROLLABLE LIST OF ALL 22+ SLOTS -->

                <ScrollViewer Grid.Row="4" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">

                    <StackPanel x:Name="StackSlotsContainer" Margin="10,6,10,10">

                        <!-- Populated dynamically via PowerShell code -->

                    </StackPanel>

                </ScrollViewer>



                <!-- 6. LEFT PANE ACTION TOOLBAR -->

                <Border Grid.Row="5" Background="#1F202B" BorderBrush="#2E3040" BorderThickness="0,1,0,0" Padding="12,10">

                    <Grid>

                        <Grid.ColumnDefinitions>

                            <ColumnDefinition Width="*"/>

                            <ColumnDefinition Width="*"/>

                        </Grid.ColumnDefinitions>

                        <Grid.RowDefinitions>

                            <RowDefinition Height="Auto"/>

                            <RowDefinition Height="Auto"/>

                        </Grid.RowDefinitions>



                        <Button x:Name="BtnCopyPowerShell" Grid.Row="0" Grid.Column="0" Margin="0,0,4,6" Background="#0C5B32" BorderBrush="#16854C">

                            <TextBlock Text="Copy Script" FontWeight="Bold" FontSize="12"/>

                        </Button>



                        <Button x:Name="BtnSaveScript" Grid.Row="0" Grid.Column="1" Margin="4,0,0,6" Background="#263A5C" BorderBrush="#3B598C">

                            <TextBlock Text="Save .ps1 File" FontWeight="SemiBold" FontSize="12"/>

                        </Button>



                        <Button x:Name="BtnCopyJson" Grid.Row="1" Grid.Column="0" Margin="0,0,4,0" Background="#2E3142" BorderBrush="#454A63">

                            <TextBlock Text="Copy JSON" FontWeight="SemiBold" FontSize="12"/>

                        </Button>



                        <Button x:Name="BtnDeployNow" Grid.Row="1" Grid.Column="1" Margin="4,0,0,0" Background="#6B2B3A" BorderBrush="#9E3E54">

                            <TextBlock Text="Deploy to Tenant" FontWeight="Bold" FontSize="12"/>

                        </Button>

                    </Grid>

                </Border>

            </Grid>



            <!-- SPLITTER -->

            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Center" VerticalAlignment="Stretch" Background="#2A2C3A"/>



            <!-- RIGHT PANE: INTERACTIVE INSPECTOR & SHAREPOINT LIVE MOCKUP -->

            <Grid Grid.Column="2" Background="#13141A">

                <Grid.RowDefinitions>

                    <RowDefinition Height="Auto"/> <!-- Interactive Inspector HUD Card -->

                    <RowDefinition Height="*"/>    <!-- Interactive Modern SharePoint Page Mockup -->

                </Grid.RowDefinitions>



                <!-- INTERACTIVE INSPECTOR HUD BANNER -->

                <Border Grid.Row="0" Margin="16,12,16,8" Background="#1D1E2A" BorderBrush="#333649" BorderThickness="1" CornerRadius="8" Padding="16,10">

                    <Grid>

                        <Grid.ColumnDefinitions>

                            <ColumnDefinition Width="Auto"/>

                            <ColumnDefinition Width="*"/>

                            <ColumnDefinition Width="Auto"/>

                        </Grid.ColumnDefinitions>



                        <!-- Color Swatch Badge -->

                        <Border Grid.Column="0" x:Name="InspectorColorBadge" Width="44" Height="44" CornerRadius="8" Background="#0078D4" BorderBrush="#FFFFFF" BorderThickness="2" Margin="0,0,16,0">

                            <Border.Effect>

                                <DropShadowEffect BlurRadius="10" ShadowDepth="0" Color="#000000" Opacity="0.5"/>

                            </Border.Effect>

                        </Border>



                        <!-- Text Details -->

                        <StackPanel Grid.Column="1" VerticalAlignment="Center">

                            <StackPanel Orientation="Horizontal" Margin="0,0,0,2">

                                <TextBlock x:Name="InspectorSlotName" Text="themePrimary" FontSize="15" FontWeight="Bold" Foreground="#FFFFFF" Margin="0,0,10,0"/>

                                <Border Background="#2B3045" CornerRadius="4" Padding="6,2" Margin="0,0,10,0">

                                    <TextBlock x:Name="InspectorHexCode" Text="#0078D4" FontFamily="Consolas" FontWeight="Bold" FontSize="12" Foreground="#61AFEF"/>

                                </Border>

                                <Border x:Name="InspectorCategoryBadge" Background="#1F3A2B" CornerRadius="4" Padding="6,2">

                                    <TextBlock x:Name="InspectorCategoryText" Text="Brand Accent" FontSize="11" Foreground="#4CD98B"/>

                                </Border>

                            </StackPanel>

                            <TextBlock x:Name="InspectorDescription" Text="Click any (i) info icon or component on the mockup below to inspect where this hex code applies." TextWrapping="Wrap" FontSize="11.5" Foreground="#B0B5C6"/>

                        </StackPanel>



                        <!-- Quick Action in HUD -->

                        <StackPanel Grid.Column="2" VerticalAlignment="Center" Margin="10,0,0,0">

                            <Button x:Name="BtnPickInspectorColor" Background="#2B2D3E" BorderBrush="#454963" Padding="10,5">

                                <TextBlock Text="Change Color" FontSize="11"/>

                            </Button>

                        </StackPanel>

                    </Grid>

                </Border>



                <!-- INTERACTIVE SHAREPOINT PAGE MOCKUP -->

                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="16,0,16,14">

                    <!-- Page Canvas (Mapped to neutralLighter or white) -->

                    <Border x:Name="mockPageCanvas" Background="#FFFFFF" BorderBrush="#D2D0CE" BorderThickness="1" CornerRadius="4" MinWidth="820" MaxWidth="1100" HorizontalAlignment="Center" Margin="0,0,0,20">

                        <Border.Effect>

                            <DropShadowEffect BlurRadius="18" ShadowDepth="3" Color="#000000" Opacity="0.35"/>

                        </Border.Effect>



                        <StackPanel>

                            <!-- 1. OFFICE 365 SUITE BAR -->

                            <Border x:Name="mockSuiteBar" Background="#0078D4" Height="48" Padding="14,0" Cursor="Hand" ToolTip="Suite Bar [themePrimary / white]">

                                <Grid>

                                    <Grid.ColumnDefinitions>

                                        <ColumnDefinition Width="Auto"/>

                                        <ColumnDefinition Width="*"/>

                                        <ColumnDefinition Width="Auto"/>

                                    </Grid.ColumnDefinitions>



                                    <!-- App Launcher & Title -->

                                    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">

                                        <!-- Clean 3x3 App Launcher Waffle -->

                                        <Border Width="20" Height="20" Background="#1665A8" CornerRadius="3" Margin="0,0,10,0" Padding="3">

                                            <UniformGrid Rows="3" Columns="3">

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                                <Border Background="#FFFFFF" Width="2.5" Height="2.5" CornerRadius="1"/>

                                            </UniformGrid>

                                        </Border>

                                        <TextBlock Text="SharePoint" FontSize="15" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,0,8,0" VerticalAlignment="Center"/>

                                        <TextBlock Text="|" FontSize="13" Foreground="#9FD1F7" Margin="0,0,8,0" VerticalAlignment="Center"/>

                                        <TextBlock Text="Contoso Hub" FontSize="13" Foreground="#E1F0FA" VerticalAlignment="Center"/>

                                        

                                        <!-- Suite Bar Info Icon (i) -->

                                        <Border x:Name="infoBtnSuiteBar" Width="18" Height="18" CornerRadius="9" Background="#004578" BorderBrush="#FFFFFF" BorderThickness="1" Cursor="Hand" Margin="10,0,0,0" VerticalAlignment="Center" ToolTip="Info: Suite Bar uses themePrimary">

                                            <TextBlock Text="i" FontSize="11" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                        </Border>

                                    </StackPanel>



                                    <!-- Suite Search Box -->

                                    <Border Grid.Column="1" HorizontalAlignment="Center" Width="360" Height="28" CornerRadius="4" Background="#FFFFFF" Padding="8,0" VerticalAlignment="Center">

                                        <Grid>

                                            <TextBlock x:Name="mockSearchPlaceholder" Text="Search this site and resources..." VerticalAlignment="Center" Foreground="#A19F9D" FontSize="12"/>

                                        </Grid>

                                    </Border>



                                    <!-- User Badges -->

                                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">

                                        <Border Width="28" Height="28" CornerRadius="14" Background="#FFFFFF">

                                            <TextBlock Text="JM" FontSize="11" FontWeight="Bold" Foreground="#0078D4" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                                        </Border>

                                    </StackPanel>

                                </Grid>

                            </Border>



                            <!-- 2. SITE HEADER & NAVIGATION BAR -->

                            <Border x:Name="mockSiteHeader" Background="#FFFFFF" BorderBrush="#EDEBE9" BorderThickness="0,0,0,1" Padding="24,12" Cursor="Hand" ToolTip="Site Header [white / neutralPrimary]">

                                <Grid>

                                    <Grid.ColumnDefinitions>

                                        <ColumnDefinition Width="Auto"/>

                                        <ColumnDefinition Width="*"/>

                                    </Grid.ColumnDefinitions>



                                    <!-- Site Logo & Title -->

                                    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">

                                        <Border Width="38" Height="38" CornerRadius="6" Background="#0078D4" Margin="0,0,12,0">

                                            <TextBlock Text="C" FontSize="18" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                                        </Border>

                                        <StackPanel VerticalAlignment="Center">

                                            <StackPanel Orientation="Horizontal">

                                                <TextBlock x:Name="mockSiteTitle" Text="Global Operations and Engineering Portal" FontSize="16" FontWeight="Bold" Foreground="#323130"/>

                                                <!-- Header Info Icon (i) -->

                                                <Border x:Name="infoBtnHeader" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" Margin="8,0,0,0" VerticalAlignment="Center" ToolTip="Info: Site Title uses neutralPrimary, canvas uses white">

                                                    <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                                </Border>

                                            </StackPanel>

                                            <TextBlock Text="Enterprise Communication Site" FontSize="11" Foreground="#605E5C"/>

                                        </StackPanel>

                                    </StackPanel>



                                    <!-- Navigation Links -->

                                    <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">

                                        <!-- Active Nav Tab -->

                                        <StackPanel Margin="0,0,18,0" Cursor="Hand">

                                            <TextBlock x:Name="mockNavActive" Text="Home Overview" FontSize="12.5" FontWeight="Bold" Foreground="#005A9E"/>

                                            <Border x:Name="mockNavUnderline" Height="3" Background="#0078D4" Margin="0,3,0,0" CornerRadius="1.5"/>

                                        </StackPanel>



                                        <TextBlock x:Name="mockNavTab1" Text="Documents" FontSize="12.5" Foreground="#605E5C" Margin="0,0,18,0" Cursor="Hand"/>

                                        <TextBlock x:Name="mockNavTab2" Text="Projects" FontSize="12.5" Foreground="#605E5C" Margin="0,0,18,0" Cursor="Hand"/>

                                        <TextBlock x:Name="mockNavTab3" Text="Team Directory" FontSize="12.5" Foreground="#605E5C" Margin="0,0,18,0" Cursor="Hand"/>



                                        <!-- Nav Info Icon (i) -->

                                        <Border x:Name="infoBtnNav" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" Margin="4,0,0,0" VerticalAlignment="Center" ToolTip="Info: Active Tab underline uses themePrimary, inactive text uses neutralSecondary">

                                            <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                        </Border>

                                    </StackPanel>

                                </Grid>

                            </Border>



                            <!-- 3. HERO / SPOTLIGHT WEB PART BANNER -->

                            <Border x:Name="mockHeroBanner" Margin="24,14,24,12" CornerRadius="6" Background="#106EBE" Padding="24,20" Cursor="Hand" ToolTip="Hero WebPart [themeDarkAlt / white / themePrimary]">

                                <Grid>

                                    <Grid.ColumnDefinitions>

                                        <ColumnDefinition Width="*"/>

                                        <ColumnDefinition Width="Auto"/>

                                    </Grid.ColumnDefinitions>



                                    <StackPanel Grid.Column="0">

                                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">

                                            <Border x:Name="mockHeroTag" Background="#71AFE5" CornerRadius="3" Padding="6,2" HorizontalAlignment="Left">

                                                <TextBlock Text="FEATURED INITIATIVE" FontSize="9.5" FontWeight="Bold" Foreground="#FFFFFF"/>

                                            </Border>

                                            <!-- Hero Banner Info Icon (i) -->

                                            <Border x:Name="infoBtnHero" Width="16" Height="16" CornerRadius="8" Background="#002D50" BorderBrush="#FFFFFF" BorderThickness="1" Cursor="Hand" Margin="8,0,0,0" VerticalAlignment="Center" ToolTip="Info: Hero background uses themeDarkAlt, Tag uses themeTertiary">

                                                <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                            </Border>

                                        </StackPanel>

                                        

                                        <TextBlock x:Name="mockHeroHeading" Text="Modern Cloud Workplace and Operations" FontSize="20" FontWeight="Bold" Foreground="#FFFFFF" Margin="0,0,0,4"/>

                                        <TextBlock Text="Empower distributed enterprise teams with unified collaboration, real-time analytics, and secure document workflows." FontSize="12.5" Foreground="#E1DFDD" MaxWidth="640" TextWrapping="Wrap" Margin="0,0,0,14"/>



                                        <StackPanel Orientation="Horizontal">

                                            <!-- Primary CTA Button -->

                                            <StackPanel Orientation="Horizontal" Margin="0,0,12,0">

                                                <Border x:Name="mockHeroBtnPrimary" Background="#0078D4" CornerRadius="3" Padding="14,7" Cursor="Hand" ToolTip="Primary Button [themePrimary / white]">

                                                    <TextBlock x:Name="mockHeroBtnPrimaryText" Text="Launch Dashboard" FontWeight="Bold" Foreground="#FFFFFF" FontSize="12"/>

                                                </Border>

                                                <!-- Primary Button Info Icon (i) -->

                                                <Border x:Name="infoBtnPrimaryBtn" Width="16" Height="16" CornerRadius="8" Background="#004578" BorderBrush="#FFFFFF" BorderThickness="1" Cursor="Hand" Margin="4,0,0,0" VerticalAlignment="Center" ToolTip="Info: Primary Action Buttons use themePrimary fill and white text">

                                                    <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                                </Border>

                                            </StackPanel>



                                            <!-- Secondary Button -->

                                            <StackPanel Orientation="Horizontal">

                                                <Border x:Name="mockHeroBtnSecondary" Background="#F3F2F1" CornerRadius="3" Padding="14,7" Cursor="Hand" ToolTip="Secondary Button [neutralLighter / neutralPrimary]">

                                                    <TextBlock Text="Documentation" FontWeight="SemiBold" Foreground="#323130" FontSize="12"/>

                                                </Border>

                                                <!-- Secondary Button Info Icon (i) -->

                                                <Border x:Name="infoBtnSecondaryBtn" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#708090" BorderThickness="1" Cursor="Hand" Margin="4,0,0,0" VerticalAlignment="Center" ToolTip="Info: Secondary Button uses neutralLighter fill and neutralPrimary text">

                                                    <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                                </Border>

                                            </StackPanel>

                                        </StackPanel>

                                    </StackPanel>



                                    <!-- Hero Metric Card -->

                                    <Border Grid.Column="1" Width="110" Height="80" Background="#2B88D8" CornerRadius="6" VerticalAlignment="Center" Padding="8">

                                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">

                                            <TextBlock Text="99.98%" FontSize="18" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center"/>

                                            <TextBlock Text="SLA Uptime" FontSize="10.5" Foreground="#C7E0F4" HorizontalAlignment="Center"/>

                                        </StackPanel>

                                    </Border>

                                </Grid>

                            </Border>



                            <!-- 4. QUICK LINKS CARDS SECTION -->

                            <StackPanel Margin="24,6,24,12">

                                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">

                                    <TextBlock x:Name="mockSectionTitle" Text="Operations and Quick Links" FontSize="14.5" FontWeight="Bold" Foreground="#201F1E"/>

                                    <!-- Quick Links Info Icon (i) -->

                                    <Border x:Name="infoBtnQuickLinks" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" Margin="8,0,0,0" VerticalAlignment="Center" ToolTip="Info: Cards use neutralLighterAlt surface, neutralLight borders, and themePrimary icons">

                                        <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                    </Border>

                                </StackPanel>

                                

                                <Grid>

                                    <Grid.ColumnDefinitions>

                                        <ColumnDefinition Width="*"/>

                                        <ColumnDefinition Width="*"/>

                                        <ColumnDefinition Width="*"/>

                                    </Grid.ColumnDefinitions>



                                    <!-- Card 1 -->

                                    <Border x:Name="mockCard1" Grid.Column="0" Margin="0,0,6,0" Background="#FAF9F8" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="4" Padding="12,10" Cursor="Hand">

                                        <StackPanel>

                                            <Border Width="28" Height="28" CornerRadius="4" Background="#EFF6FC" HorizontalAlignment="Left" Margin="0,0,0,6">

                                                <TextBlock x:Name="mockCardIcon1" Text="IT" FontSize="11" FontWeight="Bold" Foreground="#0078D4" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                                            </Border>

                                            <TextBlock x:Name="mockCardTitle1" Text="IT Governance" FontSize="12" FontWeight="Bold" Foreground="#323130" Margin="0,0,0,2"/>

                                            <TextBlock x:Name="mockCardDesc1" Text="Security baseline protocols and compliance rules." FontSize="10.5" Foreground="#605E5C" TextWrapping="Wrap"/>

                                        </StackPanel>

                                    </Border>



                                    <!-- Card 2 -->

                                    <Border x:Name="mockCard2" Grid.Column="1" Margin="3,0,3,0" Background="#FAF9F8" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="4" Padding="12,10" Cursor="Hand">

                                        <StackPanel>

                                            <Border Width="28" Height="28" CornerRadius="4" Background="#EFF6FC" HorizontalAlignment="Left" Margin="0,0,0,6">

                                                <TextBlock x:Name="mockCardIcon2" Text="KPI" FontSize="11" FontWeight="Bold" Foreground="#0078D4" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                                            </Border>

                                            <TextBlock x:Name="mockCardTitle2" Text="Operational Metrics" FontSize="12" FontWeight="Bold" Foreground="#323130" Margin="0,0,0,2"/>

                                            <TextBlock x:Name="mockCardDesc2" Text="Telemetry, cloud consumption, and team velocity." FontSize="10.5" Foreground="#605E5C" TextWrapping="Wrap"/>

                                        </StackPanel>

                                    </Border>



                                    <!-- Card 3 -->

                                    <Border x:Name="mockCard3" Grid.Column="2" Margin="6,0,0,0" Background="#FAF9F8" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="4" Padding="12,10" Cursor="Hand">

                                        <StackPanel>

                                            <Border Width="28" Height="28" CornerRadius="4" Background="#EFF6FC" HorizontalAlignment="Left" Margin="0,0,0,6">

                                                <TextBlock x:Name="mockCardIcon3" Text="DIR" FontSize="11" FontWeight="Bold" Foreground="#0078D4" HorizontalAlignment="Center" VerticalAlignment="Center"/>

                                            </Border>

                                            <TextBlock x:Name="mockCardTitle3" Text="Team Directory" FontSize="12" FontWeight="Bold" Foreground="#323130" Margin="0,0,0,2"/>

                                            <TextBlock x:Name="mockCardDesc3" Text="Organization charts, contacts, and service desk." FontSize="10.5" Foreground="#605E5C" TextWrapping="Wrap"/>

                                        </StackPanel>

                                    </Border>

                                </Grid>

                            </StackPanel>



                            <!-- 5. DOCUMENT LIBRARY LIST WEB PART -->

                            <Border x:Name="mockDocLibrary" Margin="24,4,24,12" Background="#FAF9F8" BorderBrush="#EDEBE9" BorderThickness="1" CornerRadius="4" Padding="14,12" Cursor="Hand">

                                <StackPanel>

                                    <!-- Command Bar -->

                                    <Grid Margin="0,0,0,10">

                                        <Grid.ColumnDefinitions>

                                            <ColumnDefinition Width="Auto"/>

                                            <ColumnDefinition Width="*"/>

                                            <ColumnDefinition Width="Auto"/>

                                        </Grid.ColumnDefinitions>



                                        <StackPanel Grid.Column="0" Orientation="Horizontal">

                                            <!-- + New Button (Primary) -->

                                            <Border x:Name="mockBtnPrimaryNew" Background="#0078D4" CornerRadius="3" Padding="10,4" Margin="0,0,6,0" Cursor="Hand">

                                                <TextBlock x:Name="mockBtnPrimaryNewText" Text="New Document" FontWeight="Bold" Foreground="#FFFFFF" FontSize="11.5"/>

                                            </Border>

                                            <!-- Upload (Secondary) -->

                                            <Border x:Name="mockBtnUpload" Background="#F3F2F1" BorderBrush="#D2D0CE" BorderThickness="1" CornerRadius="3" Padding="8,4" Margin="0,0,6,0" Cursor="Hand">

                                                <TextBlock Text="Upload" FontSize="11.5" Foreground="#323130" FontWeight="SemiBold"/>

                                            </Border>

                                            <!-- Sync -->

                                            <Border Background="#F3F2F1" BorderBrush="#D2D0CE" BorderThickness="1" CornerRadius="3" Padding="8,4" Cursor="Hand">

                                                <TextBlock Text="Sync" FontSize="11.5" Foreground="#323130" FontWeight="SemiBold"/>

                                            </Border>



                                            <!-- Doc Library Info Icon (i) -->

                                            <Border x:Name="infoBtnDocLib" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" Margin="6,0,0,0" VerticalAlignment="Center" ToolTip="Info: New Document uses themePrimary, selected row uses themeLighterAlt">

                                                <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                            </Border>

                                        </StackPanel>



                                        <TextBlock Grid.Column="2" Text="Documents Library (24 items)" FontSize="11" Foreground="#605E5C" VerticalAlignment="Center"/>

                                    </Grid>



                                    <!-- Table Header -->

                                    <Border BorderBrush="#EDEBE9" BorderThickness="0,0,0,1" Padding="0,0,0,5">

                                        <Grid>

                                            <Grid.ColumnDefinitions>

                                                <ColumnDefinition Width="30"/>

                                                <ColumnDefinition Width="*"/>

                                                <ColumnDefinition Width="130"/>

                                                <ColumnDefinition Width="130"/>

                                            </Grid.ColumnDefinitions>

                                            <TextBlock Grid.Column="1" x:Name="mockThName" Text="NAME" FontSize="10.5" FontWeight="Bold" Foreground="#605E5C"/>

                                            <TextBlock Grid.Column="2" x:Name="mockThModified" Text="MODIFIED" FontSize="10.5" FontWeight="Bold" Foreground="#605E5C"/>

                                            <TextBlock Grid.Column="3" x:Name="mockThAuthor" Text="AUTHOR" FontSize="10.5" FontWeight="Bold" Foreground="#605E5C"/>

                                        </Grid>

                                    </Border>



                                    <!-- Row 1 (Selected Row) -->

                                    <Border x:Name="mockRowSelected" Background="#EFF6FC" BorderBrush="#E1DFDD" BorderThickness="0,0,0,1" Padding="0,6" Cursor="Hand">

                                        <Grid>

                                            <Grid.ColumnDefinitions>

                                                <ColumnDefinition Width="30"/>

                                                <ColumnDefinition Width="*"/>

                                                <ColumnDefinition Width="130"/>

                                                <ColumnDefinition Width="130"/>

                                            </Grid.ColumnDefinitions>

                                            <TextBlock Grid.Column="0" Text="[x]" Foreground="#0078D4" FontWeight="Bold" FontSize="11" HorizontalAlignment="Center"/>

                                            <TextBlock Grid.Column="1" x:Name="mockRow1Title" Text="SharePoint-Migration-Strategy-2026.docx" FontSize="11.5" FontWeight="SemiBold" Foreground="#323130"/>

                                            <TextBlock Grid.Column="2" x:Name="mockRow1Time" Text="Just now" FontSize="10.5" Foreground="#605E5C"/>

                                            <TextBlock Grid.Column="3" Text="IT Architecture" FontSize="10.5" Foreground="#323130"/>

                                        </Grid>

                                    </Border>



                                    <!-- Row 2 (Standard Row) -->

                                    <Border x:Name="mockRowStandard" Background="Transparent" Padding="0,6" Cursor="Hand">

                                        <Grid>

                                            <Grid.ColumnDefinitions>

                                                <ColumnDefinition Width="30"/>

                                                <ColumnDefinition Width="*"/>

                                                <ColumnDefinition Width="130"/>

                                                <ColumnDefinition Width="130"/>

                                            </Grid.ColumnDefinitions>

                                            <TextBlock Grid.Column="0" Text="[ ]" Foreground="#A19F9D" FontSize="11" HorizontalAlignment="Center"/>

                                            <TextBlock Grid.Column="1" x:Name="mockRow2Title" Text="Enterprise-Cloud-Governance-Matrix.xlsx" FontSize="11.5" Foreground="#323130"/>

                                            <TextBlock Grid.Column="2" Text="Yesterday" FontSize="10.5" Foreground="#605E5C"/>

                                            <TextBlock Grid.Column="3" Text="Compliance Officer" FontSize="10.5" Foreground="#323130"/>

                                        </Grid>

                                    </Border>

                                </StackPanel>

                            </Border>



                            <!-- 6. NEWS ARTICLE & RICH TEXT WEB PART -->

                            <Border Margin="24,4,24,16" Padding="10,6" Background="Transparent">

                                <StackPanel>

                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,4">

                                        <TextBlock x:Name="mockNewsH1" Text="Corporate Modern Experience Guidelines" FontSize="15" FontWeight="Bold" Foreground="#201F1E"/>

                                        <!-- News Info Icon (i) -->

                                        <Border x:Name="infoBtnNews" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" Margin="8,0,0,0" VerticalAlignment="Center" ToolTip="Info: News H1 uses neutralDark, body uses neutralPrimary, hyperlink uses themePrimary">

                                            <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                        </Border>

                                    </StackPanel>

                                    

                                    <TextBlock x:Name="mockNewsBody" Text="SharePoint Online themes define a consistent visual identity across modern team sites, hub hierarchies, and viva connections portals. Every color slot controls specific elements from primary buttons to list borders." FontSize="11.5" Foreground="#323130" LineHeight="16" TextWrapping="Wrap" Margin="0,0,0,6"/>



                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">

                                        <TextBlock x:Name="mockLinkSample" Text="Read complete theme deployment guidelines" FontSize="11.5" FontWeight="SemiBold" Foreground="#0078D4" Cursor="Hand" Margin="0,0,14,0"/>

                                        <Border x:Name="mockDisabledTag" Background="#F3F2F1" CornerRadius="3" Padding="5,2" Margin="0,0,8,0">

                                            <TextBlock x:Name="mockDisabledText" Text="Archived" FontSize="9.5" Foreground="#A19F9D"/>

                                        </Border>

                                        <TextBlock x:Name="mockTimestamp" Text="Updated Oct 30, 2026 by Tenant Admin" FontSize="10.5" Foreground="#A19F9D"/>

                                    </StackPanel>

                                </StackPanel>

                            </Border>



                            <!-- 7. SITE FOOTER -->

                            <Border x:Name="mockFooter" Background="#F3F2F1" BorderBrush="#EDEBE9" BorderThickness="0,1,0,0" Padding="24,10" Cursor="Hand" ToolTip="Footer [neutralLighter / neutralSecondary]">

                                <Grid>

                                    <Grid.ColumnDefinitions>

                                        <ColumnDefinition Width="*"/>

                                        <ColumnDefinition Width="Auto"/>

                                    </Grid.ColumnDefinitions>

                                    <TextBlock x:Name="mockFooterText" Text="(C) 2026 Contoso Corporation. All rights reserved. | Privacy Policy | Support" FontSize="10.5" Foreground="#605E5C"/>

                                    <!-- Footer Info Icon (i) -->

                                    <Border Grid.Column="1" x:Name="infoBtnFooter" Width="16" Height="16" CornerRadius="8" Background="#E0E5F0" BorderBrush="#607090" BorderThickness="1" Cursor="Hand" ToolTip="Info: Footer background uses neutralLighter, text uses neutralSecondary">

                                        <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#203050" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Georgia"/>

                                    </Border>

                                </Grid>

                            </Border>

                        </StackPanel>

                    </Border>

                </ScrollViewer>

            </Grid>

        </Grid>



        <!-- STATUS BAR -->

        <Border Grid.Row="2" Background="#161720" BorderBrush="#252733" BorderThickness="0,1,0,0" Padding="14,6">

            <Grid>

                <Grid.ColumnDefinitions>

                    <ColumnDefinition Width="*"/>

                    <ColumnDefinition Width="Auto"/>

                </Grid.ColumnDefinitions>

                <TextBlock x:Name="TxtStatusMessage" Text="Ready. Click any (i) icon or element on the preview page to inspect." FontSize="11" Foreground="#8C92A4"/>

                <TextBlock x:Name="TxtSlotCounter" Grid.Column="1" Text="28 Theme Slots Active (22 required + 6 semantic)" FontSize="11" Foreground="#8C92A4"/>

            </Grid>

        </Border>

    </Grid>

</Window>

"@



# ---------------------------------------------------------------------------

# WINDOW INSTANTIATION & CONTROLS BINDING

# ---------------------------------------------------------------------------



try {

    $stringReader = [System.IO.StringReader]::new($xamlString)

    $xmlReader = [System.Xml.XmlReader]::Create($stringReader)

    $window = [System.Windows.Markup.XamlReader]::Load($xmlReader)

} catch {

    Write-Error "Failed to parse XAML: $_"

    return

}



# Find named controls

$comboPresets           = $window.FindName("ComboPresets")

$btnRandomize           = $window.FindName("BtnRandomize")

$btnInvertMode          = $window.FindName("BtnInvertMode")

$txtInvertMode          = $window.FindName("TxtInvertMode")

$btnFetchTenant         = $window.FindName("BtnFetchTenant")

$btnTopCopyScript       = $window.FindName("BtnTopCopyScript")



$txtThemeName           = $window.FindName("TxtThemeName")

$txtAdminUrl            = $window.FindName("TxtAdminUrl")

$txtSiteUrl             = $window.FindName("TxtSiteUrl")

$chkIsInverted          = $window.FindName("ChkIsInverted")

$chkOverwrite           = $window.FindName("ChkOverwrite")



$txtTenantId            = $window.FindName("TxtTenantId")

$txtClientId            = $window.FindName("TxtClientId")

$txtClientSecret        = $window.FindName("TxtClientSecret")



$txtQuickBaseHex        = $window.FindName("TxtQuickBaseHex")

$btnGenerateFromBase    = $window.FindName("BtnGenerateFromBase")

$txtFilterSlots         = $window.FindName("TxtFilterSlots")

$txtFilterPlaceholder   = $window.FindName("TxtFilterPlaceholder")

$stackSlotsContainer    = $window.FindName("StackSlotsContainer")



$btnCopyPowerShell      = $window.FindName("BtnCopyPowerShell")

$btnSaveScript          = $window.FindName("BtnSaveScript")

$btnCopyJson            = $window.FindName("BtnCopyJson")

$btnDeployNow           = $window.FindName("BtnDeployNow")



$inspectorColorBadge    = $window.FindName("InspectorColorBadge")

$inspectorSlotName      = $window.FindName("InspectorSlotName")

$inspectorHexCode       = $window.FindName("InspectorHexCode")

$inspectorCategoryBadge = $window.FindName("InspectorCategoryBadge")

$inspectorCategoryText  = $window.FindName("InspectorCategoryText")

$inspectorDescription   = $window.FindName("InspectorDescription")

$btnPickInspectorColor  = $window.FindName("BtnPickInspectorColor")



$txtStatusMessage       = $window.FindName("TxtStatusMessage")

$txtSlotCounter         = $window.FindName("TxtSlotCounter")



# Find Mockup Preview Elements

$mockPageCanvas         = $window.FindName("mockPageCanvas")

$mockSuiteBar           = $window.FindName("mockSuiteBar")

$mockSearchPlaceholder  = $window.FindName("mockSearchPlaceholder")

$mockSiteHeader         = $window.FindName("mockSiteHeader")

$mockSiteTitle          = $window.FindName("mockSiteTitle")

$mockNavActive          = $window.FindName("mockNavActive")

$mockNavUnderline       = $window.FindName("mockNavUnderline")

$mockNavTab1            = $window.FindName("mockNavTab1")

$mockNavTab2            = $window.FindName("mockNavTab2")

$mockNavTab3            = $window.FindName("mockNavTab3")



$mockHeroBanner         = $window.FindName("mockHeroBanner")

$mockHeroTag            = $window.FindName("mockHeroTag")

$mockHeroHeading        = $window.FindName("mockHeroHeading")

$mockHeroBtnPrimary     = $window.FindName("mockHeroBtnPrimary")

$mockHeroBtnPrimaryText = $window.FindName("mockHeroBtnPrimaryText")

$mockHeroBtnSecondary   = $window.FindName("mockHeroBtnSecondary")



$mockSectionTitle       = $window.FindName("mockSectionTitle")

$mockCard1              = $window.FindName("mockCard1")

$mockCard2              = $window.FindName("mockCard2")

$mockCard3              = $window.FindName("mockCard3")

$mockCardIcon1          = $window.FindName("mockCardIcon1")

$mockCardIcon2          = $window.FindName("mockCardIcon2")

$mockCardIcon3          = $window.FindName("mockCardIcon3")

$mockCardTitle1         = $window.FindName("mockCardTitle1")

$mockCardTitle2         = $window.FindName("mockCardTitle2")

$mockCardTitle3         = $window.FindName("mockCardTitle3")

$mockCardDesc1          = $window.FindName("mockCardDesc1")

$mockCardDesc2          = $window.FindName("mockCardDesc2")

$mockCardDesc3          = $window.FindName("mockCardDesc3")



$mockDocLibrary         = $window.FindName("mockDocLibrary")

$mockBtnPrimaryNew      = $window.FindName("mockBtnPrimaryNew")

$mockBtnPrimaryNewText  = $window.FindName("mockBtnPrimaryNewText")

$mockBtnUpload          = $window.FindName("mockBtnUpload")

$mockThName             = $window.FindName("mockThName")

$mockThModified         = $window.FindName("mockThModified")

$mockThAuthor           = $window.FindName("mockThAuthor")

$mockRowSelected        = $window.FindName("mockRowSelected")

$mockRow1Title          = $window.FindName("mockRow1Title")

$mockRow1Time           = $window.FindName("mockRow1Time")

$mockRowStandard        = $window.FindName("mockRowStandard")

$mockRow2Title          = $window.FindName("mockRow2Title")



$mockNewsH1             = $window.FindName("mockNewsH1")

$mockNewsBody           = $window.FindName("mockNewsBody")

$mockLinkSample         = $window.FindName("mockLinkSample")

$mockDisabledTag        = $window.FindName("mockDisabledTag")

$mockDisabledText       = $window.FindName("mockDisabledText")

$mockTimestamp          = $window.FindName("mockTimestamp")

$mockFooter             = $window.FindName("mockFooter")

$mockFooterText         = $window.FindName("mockFooterText")



# Find Info (i) Buttons

$infoBtnSuiteBar        = $window.FindName("infoBtnSuiteBar")

$infoBtnHeader          = $window.FindName("infoBtnHeader")

$infoBtnNav             = $window.FindName("infoBtnNav")

$infoBtnHero            = $window.FindName("infoBtnHero")

$infoBtnPrimaryBtn      = $window.FindName("infoBtnPrimaryBtn")

$infoBtnSecondaryBtn    = $window.FindName("infoBtnSecondaryBtn")

$infoBtnQuickLinks      = $window.FindName("infoBtnQuickLinks")

$infoBtnDocLib          = $window.FindName("infoBtnDocLib")

$infoBtnNews            = $window.FindName("infoBtnNews")

$infoBtnFooter          = $window.FindName("infoBtnFooter")



# Internal state tracking

$script:CurrentPalette = [ordered]@{}

$script:SlotControls = @{}

$script:CurrentlyInspectedSlot = "themePrimary"

$script:UpdatingUI = $false



# ---------------------------------------------------------------------------

# DYNAMIC SLOT FORM GENERATION

# ---------------------------------------------------------------------------



function Build-SlotFormRows() {

    $stackSlotsContainer.Children.Clear()

    $script:SlotControls.Clear()



    $currentCategory = ""



    foreach ($slotKey in $script:SlotMetadata.Keys) {

        $meta = $script:SlotMetadata[$slotKey]



        # Insert Category Header if changed

        if ($meta.Category -ne $currentCategory) {

            $currentCategory = $meta.Category

            $catBorder = [System.Windows.Controls.Border]::new()

            $catBorder.Margin = [System.Windows.Thickness]::new(0, 10, 0, 4)

            $catBorder.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)

            $catBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 34, 36, 48))

            $catBorder.CornerRadius = [System.Windows.CornerRadius]::new(4)



            $catTitle = [System.Windows.Controls.TextBlock]::new()

            $catTitle.Text = $currentCategory.ToUpper()

            $catTitle.FontSize = 10.5

            $catTitle.FontWeight = [System.Windows.FontWeights]::Bold

            $catTitle.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 140, 146, 170))



            $catBorder.Child = $catTitle

            $stackSlotsContainer.Children.Add($catBorder) | Out-Null

        }



        # Row Border Container

        $rowBorder = [System.Windows.Controls.Border]::new()

        $rowBorder.Padding = [System.Windows.Thickness]::new(8, 5, 8, 5)

        $rowBorder.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        $rowBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 27, 28, 38))

        $rowBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 42, 44, 58))

        $rowBorder.BorderThickness = [System.Windows.Thickness]::new(1)

        $rowBorder.CornerRadius = [System.Windows.CornerRadius]::new(4)

        $rowBorder.Cursor = [System.Windows.Input.Cursors]::Hand

        $rowBorder.Tag = $slotKey



        $grid = [System.Windows.Controls.Grid]::new()

        $col0 = [System.Windows.Controls.ColumnDefinition]::new() # Color Swatch

        $col0.Width = [System.Windows.GridLength]::new(34, [System.Windows.GridUnitType]::Pixel)

        $col1 = [System.Windows.Controls.ColumnDefinition]::new() # Slot Name & Role

        $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)

        $col2 = [System.Windows.Controls.ColumnDefinition]::new() # Hex Input

        $col2.Width = [System.Windows.GridLength]::new(84, [System.Windows.GridUnitType]::Pixel)



        $grid.ColumnDefinitions.Add($col0)

        $grid.ColumnDefinitions.Add($col1)

        $grid.ColumnDefinitions.Add($col2)



        # 1. Color Swatch Button

        $swatchBtn = [System.Windows.Controls.Button]::new()

        $swatchBtn.Width = 26

        $swatchBtn.Height = 26

        $swatchBtn.BorderBrush = [System.Windows.Media.Brushes]::White

        $swatchBtn.BorderThickness = [System.Windows.Thickness]::new(1)

        $swatchBtn.Padding = [System.Windows.Thickness]::new(0)

        $swatchBtn.Tag = $slotKey

        $swatchBtn.ToolTip = "Click to pick a color for $slotKey"

        [System.Windows.Controls.Grid]::SetColumn($swatchBtn, 0)

        $grid.Children.Add($swatchBtn) | Out-Null



        # 2. Name & Summary

        $infoStack = [System.Windows.Controls.StackPanel]::new()

        $infoStack.Margin = [System.Windows.Thickness]::new(8, 0, 8, 0)

        $infoStack.VerticalAlignment = [System.Windows.VerticalAlignment]::Center



        $slotNameTxt = [System.Windows.Controls.TextBlock]::new()

        $slotNameTxt.Text = $slotKey

        $slotNameTxt.FontSize = 12

        $slotNameTxt.FontWeight = [System.Windows.FontWeights]::SemiBold

        $slotNameTxt.Foreground = [System.Windows.Media.Brushes]::White



        $slotRoleTxt = [System.Windows.Controls.TextBlock]::new()

        $slotRoleTxt.Text = $meta.Role.Substring(0, [Math]::Min(55, $meta.Role.Length)) + "..."

        $slotRoleTxt.FontSize = 10

        $slotRoleTxt.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 140, 145, 165))

        $slotRoleTxt.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis



        $infoStack.Children.Add($slotNameTxt) | Out-Null

        $infoStack.Children.Add($slotRoleTxt) | Out-Null

        [System.Windows.Controls.Grid]::SetColumn($infoStack, 1)

        $grid.Children.Add($infoStack) | Out-Null



        # 3. Hex Input Box

        $hexTxt = [System.Windows.Controls.TextBox]::new()

        $hexTxt.Height = 24

        $hexTxt.FontSize = 11

        $hexTxt.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")

        $hexTxt.MaxLength = 7

        $hexTxt.Tag = $slotKey

        $hexTxt.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        [System.Windows.Controls.Grid]::SetColumn($hexTxt, 2)

        $grid.Children.Add($hexTxt) | Out-Null



        $rowBorder.Child = $grid

        $stackSlotsContainer.Children.Add($rowBorder) | Out-Null



        # Store references for fast updates

        $script:SlotControls[$slotKey] = @{

            RowBorder    = $rowBorder

            SwatchBtn    = $swatchBtn

            HexTextBox   = $hexTxt

            SlotNameText = $slotNameTxt

        }



        # Event: Hex Input Changed

        $hexTxt.Add_TextChanged({

            param($sender, $e)

            if ($script:UpdatingUI) { return }

            $key = $sender.Tag

            $val = $sender.Text.Trim()

            if ($val -match '^#?[0-9a-fA-F]{6}$') {

                if (-not $val.StartsWith('#')) { $val = "#$val" }

                $script:CurrentPalette[$key] = $val.ToLower()

                Update-PreviewAndSwatches -ExcludeTextBox $key

                Update-Inspector $key

            }

        })



        # Event: Swatch Clicked (Open ColorPicker)

        $swatchBtn.Add_Click({

            param($sender, $e)

            $key = $sender.Tag

            Show-ColorPickerDialog $key

        })



        # Event: Row Mouse Enter (Highlight & Inspect)

        $rowBorder.Add_MouseEnter({

            param($sender, $e)

            $key = $sender.Tag

            Inspect-SlotFromForm $key

        })



        # Event: Row Click (Select & Highlight)

        $rowBorder.Add_MouseLeftButtonDown({

            param($sender, $e)

            $key = $sender.Tag

            Inspect-SlotFromForm $key

        })

    }

}



# ---------------------------------------------------------------------------

# COLOR PICKER DIALOG HANDLER

# ---------------------------------------------------------------------------



function Show-ColorPickerDialog([string]$slotKey) {

    $currentHex = $script:CurrentPalette[$slotKey]

    $c = Convert-HexToDrawingColor $currentHex



    $dialog = [System.Windows.Forms.ColorDialog]::new()

    $dialog.AllowFullOpen = $true

    $dialog.AnyColor = $true

    $dialog.Color = $c



    try {

        $customColors = @()

        if ($script:CurrentPalette.Contains("themePrimary")) {

            $customColors += [System.Drawing.ColorTranslator]::ToOle((Convert-HexToDrawingColor $script:CurrentPalette["themePrimary"]))

        }

        if ($customColors.Count -gt 0) {

            $dialog.CustomColors = $customColors

        }

    } catch {}



    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

        $newHex = Convert-DrawingColorToHex $dialog.Color

        $script:CurrentPalette[$slotKey] = $newHex

        if ($script:SlotControls.Contains($slotKey)) {

            $script:SlotControls[$slotKey].HexTextBox.Text = $newHex

        }

        Update-PreviewAndSwatches

        Update-Inspector $slotKey

        $txtStatusMessage.Text = "Updated $slotKey to $newHex."

    }

}



# ---------------------------------------------------------------------------

# INSPECTION ENGINE (BIDIRECTIONAL)

# ---------------------------------------------------------------------------



function Update-Inspector([string]$slotKey) {

    $script:CurrentlyInspectedSlot = $slotKey

    if (-not $script:SlotMetadata.Contains($slotKey)) { return }

    $meta = $script:SlotMetadata[$slotKey]

    $hex = if ($script:CurrentPalette.Contains($slotKey)) { $script:CurrentPalette[$slotKey] } else { "#0078d4" }



    $inspectorSlotName.Text = $slotKey

    $inspectorHexCode.Text = $hex.ToUpper()

    $inspectorCategoryText.Text = $meta.Category

    $inspectorDescription.Text = $meta.Role

    $inspectorColorBadge.Background = Convert-HexToMediaBrush $hex



    # Category badge styling

    if ($meta.Category -like "*Brand*") {

        $inspectorCategoryBadge.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 20, 60, 40))

        $inspectorCategoryText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 80, 220, 140))

    } else {

        $inspectorCategoryBadge.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 40, 50, 75))

        $inspectorCategoryText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 140, 180, 240))

    }



    # Pulse highlight in row list

    foreach ($k in $script:SlotControls.Keys) {

        $row = $script:SlotControls[$k].RowBorder

        if ($k -eq $slotKey) {

            $row.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 0, 150, 255))

            $row.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 36, 42, 60))

        } else {

            $row.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 42, 44, 58))

            $row.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 27, 28, 38))

        }

    }

}



function Inspect-SlotFromForm([string]$slotKey) {

    Update-Inspector $slotKey

    Highlight-PreviewElementsForSlot $slotKey

}



$script:InspectFromPreviewElement = {

    param([string]$slotKey, [string]$elementFriendlyName)

    Update-Inspector $slotKey

    $txtStatusMessage.Text = "Inspecting '$elementFriendlyName' (Slot: $slotKey)."



    if ($script:SlotControls.Contains($slotKey)) {

        $row = $script:SlotControls[$slotKey].RowBorder

        $row.BringIntoView()

    }

}



function Inspect-FromPreviewElement([string]$slotKey, [string]$elementFriendlyName) {

    & $script:InspectFromPreviewElement $slotKey $elementFriendlyName

}



function Highlight-PreviewElementsForSlot([string]$slotKey) {

    if (-not $script:SlotMetadata.Contains($slotKey)) { return }

    $elements = $script:SlotMetadata[$slotKey].UIElements



    Reset-PreviewBorders



    foreach ($elem in $elements) {

        $ctrl = $window.FindName($elem)

        if ($null -ne $ctrl) {

            if ($ctrl -is [System.Windows.Controls.Border]) {

                $ctrl.BorderBrush = [System.Windows.Media.Brushes]::OrangeRed

                $ctrl.BorderThickness = [System.Windows.Thickness]::new(2)

            }

        }

    }

}



function Reset-PreviewBorders() {

    $nl = Convert-HexToMediaBrush (Get-SlotColor "neutralLight" "#EDEBE9")

    $mockSiteHeader.BorderBrush = $nl

    $mockCard1.BorderBrush = $nl

    $mockCard2.BorderBrush = $nl

    $mockCard3.BorderBrush = $nl

    $mockDocLibrary.BorderBrush = $nl

    $mockFooter.BorderBrush = $nl



    $mockCard1.BorderThickness = [System.Windows.Thickness]::new(1)

    $mockCard2.BorderThickness = [System.Windows.Thickness]::new(1)

    $mockCard3.BorderThickness = [System.Windows.Thickness]::new(1)

    $mockDocLibrary.BorderThickness = [System.Windows.Thickness]::new(1)

    $mockFooter.BorderThickness = [System.Windows.Thickness]::new(0, 1, 0, 0)

}



function Get-SlotColor([string]$key, [string]$fallback = "#0078D4") {

    if ($script:CurrentPalette.Contains($key)) {

        return $script:CurrentPalette[$key]

    }

    return $fallback

}



# ---------------------------------------------------------------------------

# REAL-TIME SHAREPOINT PREVIEW UPDATER

# ---------------------------------------------------------------------------



function Update-PreviewAndSwatches($ExcludeTextBox = $null) {

    $script:UpdatingUI = $true

    try {

        # Update swatches and textboxes in form

        foreach ($k in $script:CurrentPalette.Keys) {

            if ($script:SlotControls.Contains($k)) {

                $hex = $script:CurrentPalette[$k]

                $ctrl = $script:SlotControls[$k]

                $brush = Convert-HexToMediaBrush $hex

                $ctrl.SwatchBtn.Background = $brush



                if ($ExcludeTextBox -ne $k) {

                    $ctrl.HexTextBox.Text = $hex

                }

            }

        }



        # Retrieve brushes for major slots

        $pPrimary        = Convert-HexToMediaBrush (Get-SlotColor "themePrimary")

        $pLighterAlt     = Convert-HexToMediaBrush (Get-SlotColor "themeLighterAlt")

        $pLighter        = Convert-HexToMediaBrush (Get-SlotColor "themeLighter")

        $pLight          = Convert-HexToMediaBrush (Get-SlotColor "themeLight")

        $pTertiary       = Convert-HexToMediaBrush (Get-SlotColor "themeTertiary")

        $pSecondary      = Convert-HexToMediaBrush (Get-SlotColor "themeSecondary")

        $pDarkAlt        = Convert-HexToMediaBrush (Get-SlotColor "themeDarkAlt")

        $pDark           = Convert-HexToMediaBrush (Get-SlotColor "themeDark")

        $pDarker         = Convert-HexToMediaBrush (Get-SlotColor "themeDarker")



        $nLighterAlt     = Convert-HexToMediaBrush (Get-SlotColor "neutralLighterAlt")

        $nLighter        = Convert-HexToMediaBrush (Get-SlotColor "neutralLighter")

        $nLight          = Convert-HexToMediaBrush (Get-SlotColor "neutralLight")

        $nQuatAlt        = Convert-HexToMediaBrush (Get-SlotColor "neutralQuaternaryAlt")

        $nQuat           = Convert-HexToMediaBrush (Get-SlotColor "neutralQuaternary")

        $nTertAlt        = Convert-HexToMediaBrush (Get-SlotColor "neutralTertiaryAlt")

        $nTert           = Convert-HexToMediaBrush (Get-SlotColor "neutralTertiary")

        $nSec            = Convert-HexToMediaBrush (Get-SlotColor "neutralSecondary")

        $nPrimAlt        = Convert-HexToMediaBrush (Get-SlotColor "neutralPrimaryAlt")

        $nPrim           = Convert-HexToMediaBrush (Get-SlotColor "neutralPrimary")

        $nDark           = Convert-HexToMediaBrush (Get-SlotColor "neutralDark")

        $cBlack          = Convert-HexToMediaBrush (Get-SlotColor "black")

        $cWhite          = Convert-HexToMediaBrush (Get-SlotColor "white")



        $isInverted = $chkIsInverted.IsChecked -eq $true



        # 1. Page Background Canvas

        $mockPageCanvas.Background = $cWhite

        $mockPageCanvas.BorderBrush = $nQuat



        # 2. Suite Bar

        $mockSuiteBar.Background = $pPrimary



        # 3. Site Header & Nav

        $mockSiteHeader.Background = $cWhite

        $mockSiteHeader.BorderBrush = $nLight

        $mockSiteTitle.Foreground = $nPrim

        $mockNavActive.Foreground = if ($isInverted) { $pLighter } else { $pDark }

        $mockNavUnderline.Background = $pPrimary

        $mockNavTab1.Foreground = $nSec

        $mockNavTab2.Foreground = $nSec

        $mockNavTab3.Foreground = $nSec



        # 4. Hero Banner

        $mockHeroBanner.Background = $pDarkAlt

        $mockHeroTag.Background = $pTertiary

        $mockHeroHeading.Foreground = $cWhite

        $mockHeroBtnPrimary.Background = $pPrimary

        $mockHeroBtnPrimaryText.Foreground = $cWhite

        $mockHeroBtnSecondary.Background = $nLighter



        # 5. Quick Links Cards

        $mockSectionTitle.Foreground = $nDark

        $cardBg = $nLighterAlt

        $cardBorder = $nLight



        $mockCard1.Background = $cardBg

        $mockCard1.BorderBrush = $cardBorder

        $mockCardTitle1.Foreground = $nPrim

        $mockCardDesc1.Foreground = $nSec

        $mockCardIcon1.Foreground = $pPrimary



        $mockCard2.Background = $cardBg

        $mockCard2.BorderBrush = $cardBorder

        $mockCardTitle2.Foreground = $nPrim

        $mockCardDesc2.Foreground = $nSec

        $mockCardIcon2.Foreground = $pPrimary



        $mockCard3.Background = $cardBg

        $mockCard3.BorderBrush = $cardBorder

        $mockCardTitle3.Foreground = $nPrim

        $mockCardDesc3.Foreground = $nSec

        $mockCardIcon3.Foreground = $pPrimary



        # 6. Document Library

        $mockDocLibrary.Background = $cardBg

        $mockDocLibrary.BorderBrush = $cardBorder

        $mockBtnPrimaryNew.Background = $pPrimary

        $mockBtnPrimaryNewText.Foreground = $cWhite

        $mockBtnUpload.Background = $nLighter

        $mockBtnUpload.BorderBrush = $nQuat

        $mockThName.Foreground = $nSec

        $mockThModified.Foreground = $nSec

        $mockThAuthor.Foreground = $nSec

        $mockRowSelected.Background = $pLighterAlt

        $mockRowSelected.BorderBrush = $nQuatAlt

        $mockRow1Title.Foreground = $nPrim

        $mockRow1Time.Foreground = $nSec

        $mockRow2Title.Foreground = $nPrim



        # 7. News & Content

        $mockNewsH1.Foreground = $nDark

        $mockNewsBody.Foreground = $nPrim

        $mockLinkSample.Foreground = $pPrimary

        $mockDisabledTag.Background = $nLighter

        $mockDisabledText.Foreground = $nTert

        $mockTimestamp.Foreground = $nTert



        # 8. Footer

        $mockFooter.Background = $nLighter

        $mockFooter.BorderBrush = $nLight

        $mockFooterText.Foreground = $nSec



        # Update Inspector swatch

        if ($script:CurrentPalette.Contains($script:CurrentlyInspectedSlot)) {

            $inspectorColorBadge.Background = Convert-HexToMediaBrush $script:CurrentPalette[$script:CurrentlyInspectedSlot]

            $inspectorHexCode.Text = $script:CurrentPalette[$script:CurrentlyInspectedSlot].ToUpper()

        }



    } finally {

        $script:UpdatingUI = $false

    }

}



# ---------------------------------------------------------------------------

# PREVIEW COMPONENT & INFO (i) BUTTON INTERACTIVITY

# ---------------------------------------------------------------------------



function Register-PreviewInteractivity() {

    $interactiveMappings = @(

        @{ Control = $mockSuiteBar;           Slot = "themePrimary";      Name = "Office 365 Suite Bar" }

        @{ Control = $infoBtnSuiteBar;        Slot = "themePrimary";      Name = "Suite Bar [themePrimary]" }

        @{ Control = $mockSiteHeader;         Slot = "white";             Name = "Site Header Surface" }

        @{ Control = $mockSiteTitle;          Slot = "neutralPrimary";    Name = "Site Title Text" }

        @{ Control = $infoBtnHeader;          Slot = "neutralPrimary";    Name = "Site Title [neutralPrimary]" }

        @{ Control = $mockNavActive;          Slot = "themeDark";         Name = "Active Navigation Tab Text" }

        @{ Control = $mockNavUnderline;       Slot = "themePrimary";      Name = "Active Navigation Tab Underline" }

        @{ Control = $mockNavTab1;            Slot = "neutralSecondary";  Name = "Inactive Navigation Tab Text" }

        @{ Control = $infoBtnNav;             Slot = "themePrimary";      Name = "Active Nav Tab [themePrimary]" }

        @{ Control = $mockHeroBanner;         Slot = "themeDarkAlt";      Name = "Hero WebPart Banner Background" }

        @{ Control = $infoBtnHero;            Slot = "themeDarkAlt";      Name = "Hero Banner Background [themeDarkAlt]" }

        @{ Control = $mockHeroTag;            Slot = "themeTertiary";     Name = "Hero WebPart Tag Pill" }

        @{ Control = $mockHeroHeading;        Slot = "white";             Name = "Hero Banner Heading" }

        @{ Control = $mockHeroBtnPrimary;     Slot = "themePrimary";      Name = "Primary Call-to-Action Button" }

        @{ Control = $infoBtnPrimaryBtn;      Slot = "themePrimary";      Name = "Primary Button [themePrimary]" }

        @{ Control = $mockHeroBtnSecondary;   Slot = "neutralLighter";    Name = "Secondary Button Background" }

        @{ Control = $infoBtnSecondaryBtn;    Slot = "neutralLighter";    Name = "Secondary Button [neutralLighter]" }

        @{ Control = $mockSectionTitle;       Slot = "neutralDark";       Name = "Section H2 Header" }

        @{ Control = $infoBtnQuickLinks;      Slot = "neutralLighterAlt"; Name = "Quick Links Container [neutralLighterAlt]" }

        @{ Control = $mockCard1;              Slot = "neutralLighterAlt"; Name = "Quick Links Card Container" }

        @{ Control = $mockCardTitle1;         Slot = "neutralPrimary";    Name = "Card Title Heading" }

        @{ Control = $mockCardDesc1;          Slot = "neutralSecondary";  Name = "Card Subtitle / Description" }

        @{ Control = $mockCardIcon1;          Slot = "themePrimary";      Name = "Card Accent Icon" }

        @{ Control = $mockDocLibrary;         Slot = "neutralLighterAlt"; Name = "Document Library Canvas" }

        @{ Control = $infoBtnDocLib;          Slot = "themePrimary";      Name = "Document Library Actions [themePrimary]" }

        @{ Control = $mockBtnPrimaryNew;      Slot = "themePrimary";      Name = "Command Bar Primary Action (New)" }

        @{ Control = $mockRowSelected;        Slot = "themeLighterAlt";   Name = "Selected Table Row Highlight" }

        @{ Control = $mockRow1Title;          Slot = "neutralPrimary";    Name = "Document Title Label" }

        @{ Control = $mockRow1Time;           Slot = "neutralSecondary";  Name = "Timestamp Metadata" }

        @{ Control = $mockNewsH1;             Slot = "neutralDark";       Name = "Article Title (H1)" }

        @{ Control = $infoBtnNews;            Slot = "neutralDark";       Name = "Article Heading [neutralDark]" }

        @{ Control = $mockNewsBody;           Slot = "neutralPrimary";    Name = "Article Body Paragraph Text" }

        @{ Control = $mockLinkSample;         Slot = "themePrimary";      Name = "Inline Hyperlink Text" }

        @{ Control = $mockDisabledTag;        Slot = "neutralLighter";    Name = "Disabled Chip Background" }

        @{ Control = $mockDisabledText;       Slot = "neutralTertiary";   Name = "Disabled Label Text" }

        @{ Control = $mockFooter;             Slot = "neutralLighter";    Name = "Site Footer Bar" }

        @{ Control = $infoBtnFooter;          Slot = "neutralLighter";    Name = "Footer Background [neutralLighter]" }

        @{ Control = $mockFooterText;         Slot = "neutralSecondary";  Name = "Footer Copyright & Secondary Links" }

    )



    $inspectHandler = {

        param($sender, $e)

        if ($null -ne $sender -and $null -ne $sender.Tag) {

            $tag = $sender.Tag

            & $script:InspectFromPreviewElement $tag.Slot $tag.Name

        }

    }



    foreach ($item in $interactiveMappings) {

        $ctrl = $item.Control

        if ($null -ne $ctrl) {

            $ctrl.Tag = @{ Slot = $item.Slot; Name = $item.Name }

            $ctrl.Add_MouseEnter($inspectHandler)

            $ctrl.Add_MouseLeftButtonDown($inspectHandler)

        }

    }

}



# ---------------------------------------------------------------------------

# LOAD PRESETS & EVENT HANDLERS

# ---------------------------------------------------------------------------



function Load-PaletteDictionary([hashtable]$paletteDict, [bool]$isInverted) {

    $script:CurrentPalette.Clear()

    foreach ($k in $paletteDict.Keys) {

        $script:CurrentPalette[$k] = $paletteDict[$k].ToLower()

    }

    $chkIsInverted.IsChecked = $isInverted

    if ($script:CurrentPalette.Contains("themePrimary")) {

        $txtQuickBaseHex.Text = $script:CurrentPalette["themePrimary"]

    }

    Update-PreviewAndSwatches

    Update-Inspector "themePrimary"

}



# Initialize Presets Combobox

foreach ($presetName in $script:Presets.Keys) {

    $comboPresets.Items.Add($presetName) | Out-Null

}

$comboPresets.SelectedIndex = 0



$comboPresets.Add_SelectionChanged({

    $selected = $comboPresets.SelectedItem

    if ($null -ne $selected -and $script:Presets.Contains($selected)) {

        $p = $script:Presets[$selected]

        $cleanPresetName = $selected -replace '[#%&*{}\\:<>\?\/+\|\"\(\)\[\]]', ''

        $txtThemeName.Text = "Contoso $cleanPresetName"

        Load-PaletteDictionary $p.Slots $p.IsInverted

        $txtStatusMessage.Text = "Loaded preset '$selected'."

    }

})



# Randomize Scheme Button

$btnRandomize.Add_Click({

    $inv = $chkIsInverted.IsChecked -eq $true

    $randomPalette = Generate-RandomScheme $inv

    $primary = $randomPalette["themePrimary"]

    $txtQuickBaseHex.Text = $primary

    $cleanSuffix = $primary.TrimStart('#').ToUpper()

    $txtThemeName.Text = "Contoso Custom $cleanSuffix"

    Load-PaletteDictionary $randomPalette $inv

    $txtStatusMessage.Text = "Generated random color scheme ($primary)."

})



# Invert / Dark Mode Toggle Button

$btnInvertMode.Add_Click({

    $chkIsInverted.IsChecked = -not ($chkIsInverted.IsChecked -eq $true)

    $inv = $chkIsInverted.IsChecked -eq $true

    $baseHex = if ($script:CurrentPalette.Contains("themePrimary")) { $script:CurrentPalette["themePrimary"] } else { "#0078d4" }

    $newPalette = Generate-PaletteFromPrimary $baseHex $inv

    Load-PaletteDictionary $newPalette $inv

    $txtStatusMessage.Text = "Toggled theme inversion. Dark Mode: $inv."

})



$chkIsInverted.Add_Checked({

    $txtInvertMode.Text = "Switch to Light Mode"

    Update-PreviewAndSwatches

})

$chkIsInverted.Add_Unchecked({

    $txtInvertMode.Text = "Switch to Dark Mode"

    Update-PreviewAndSwatches

})



# Auto-synchronize all 22 slots when Base Primary Hex changes

$txtQuickBaseHex.Add_TextChanged({

    if ($script:UpdatingUI) { return }

    $base = $txtQuickBaseHex.Text.Trim()

    if ($base -match '^#?[0-9a-fA-F]{6}$') {

        if (-not $base.StartsWith('#')) { $base = "#$base" }

        $inv = $chkIsInverted.IsChecked -eq $true

        $newPalette = Generate-PaletteFromPrimary $base $inv

        Load-PaletteDictionary $newPalette $inv

        $txtStatusMessage.Text = "Synchronized all 22 theme slots with base color $base."

    }

})



# Generate From Single Base Color Button Click

$btnGenerateFromBase.Add_Click({

    $base = $txtQuickBaseHex.Text.Trim()

    if ($base -match '^#?[0-9a-fA-F]{6}$') {

        if (-not $base.StartsWith('#')) { $base = "#$base" }

        $inv = $chkIsInverted.IsChecked -eq $true

        $newPalette = Generate-PaletteFromPrimary $base $inv

        Load-PaletteDictionary $newPalette $inv

        $txtStatusMessage.Text = "Generated full 22-slot palette from base primary $base."

    } else {

        [System.Windows.MessageBox]::Show("Please enter a valid 6-character hex code (e.g. #0078D4).", "Invalid Hex Code", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)

    }

})



# Change Inspector Color Button

$btnPickInspectorColor.Add_Click({

    if (-not [string]::IsNullOrEmpty($script:CurrentlyInspectedSlot)) {

        Show-ColorPickerDialog $script:CurrentlyInspectedSlot

    }

})



# Filter Slots Real-Time

$txtFilterSlots.Add_TextChanged({

    $filter = $txtFilterSlots.Text.Trim().ToLower()

    $txtFilterPlaceholder.Visibility = if ([string]::IsNullOrEmpty($filter)) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }



    foreach ($k in $script:SlotControls.Keys) {

        $meta = $script:SlotMetadata[$k]

        $match = ($k.ToLower().Contains($filter)) -or ($meta.DisplayName.ToLower().Contains($filter)) -or ($meta.Role.ToLower().Contains($filter)) -or ($meta.Category.ToLower().Contains($filter))

        $script:SlotControls[$k].RowBorder.Visibility = if ($match) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    }

})



# ---------------------------------------------------------------------------

# SCRIPT GENERATION & DEPLOYMENT UTILITIES

# ---------------------------------------------------------------------------



function Generate-PowerShellScript() {

    $rawThemeName = $txtThemeName.Text.Trim()
    if ([string]::IsNullOrEmpty($rawThemeName)) { $rawThemeName = "Contoso Custom Theme" }
    
    # Strict Sanitization: Theme Name cannot contain shell escapes or SPO forbidden characters
    $themeName = ($rawThemeName -replace '[#%&*{}\\:<>\?\/+\|\"\(\)\[\]''`;\$]', '') -replace '\s+', ' '
    $themeName = $themeName.Trim()
    if ([string]::IsNullOrEmpty($themeName)) { $themeName = "Contoso Custom Theme" }

    $rawAdminUrl  = $txtAdminUrl.Text.Trim()
    $adminUrl     = ($rawAdminUrl -replace '["''`;\$\s]', '').Trim()

    $rawSiteUrl   = $txtSiteUrl.Text.Trim()
    $siteUrl      = ($rawSiteUrl -split '[#\?]')[0].Trim().TrimEnd('/')
    $siteUrl      = ($siteUrl -replace '["''`;\$\s]', '').Trim()

    $inv          = if ($chkIsInverted.IsChecked) { '$true' } else { '$false' }

    $rawTenantId  = $txtTenantId.Text.Trim()
    $tenantId     = ($rawTenantId -replace '[^a-zA-Z0-9\-]', '').Trim()

    $rawClientId  = $txtClientId.Text.Trim()
    $clientId     = ($rawClientId -replace '[^a-zA-Z0-9\-]', '').Trim()

    $clientSecret = if ($null -ne $txtClientSecret) { $txtClientSecret.Password.Trim() } else { "" }



    $sb = [System.Text.StringBuilder]::new()

    $sb.AppendLine("<#") | Out-Null

    $sb.AppendLine(".SYNOPSIS") | Out-Null

    $sb.AppendLine("    Registers and applies the '$themeName' custom theme in SharePoint Online.") | Out-Null

    $sb.AppendLine(".DESCRIPTION") | Out-Null

    $sb.AppendLine("    Generated by SharePoint Online Theme Studio.") | Out-Null

    $sb.AppendLine("    Requires SharePoint Online Management Shell (Microsoft.Online.SharePoint.PowerShell).") | Out-Null

    $sb.AppendLine("#>") | Out-Null

    $sb.AppendLine("") | Out-Null
    $escapedAdmin  = $adminUrl -replace "'", "''"
    $escapedSite   = $siteUrl -replace "'", "''"
    $escapedName   = $themeName -replace "'", "''"
    $escapedTenant = $tenantId -replace "'", "''"
    $escapedClient = $clientId -replace "'", "''"
    $escapedSecret = $clientSecret -replace "'", "''"

    $sb.AppendLine("`$AdminCenterUrl = '$escapedAdmin'") | Out-Null
    $sb.AppendLine("`$TargetSiteUrl  = '$escapedSite'") | Out-Null
    $sb.AppendLine("`$ThemeName      = '$escapedName'") | Out-Null
    
    if (-not [string]::IsNullOrWhiteSpace($tenantId)) {
        $sb.AppendLine("`$TenantId       = '$escapedTenant'") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($clientId)) {
        $sb.AppendLine("`$ClientId       = '$escapedClient'") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($clientSecret)) {
        $sb.AppendLine("`$ClientSecret   = '$escapedSecret'") | Out-Null
    }



    $sb.AppendLine("") | Out-Null

    $sb.AppendLine("`$ThemePalette = @{") | Out-Null



    $officialSlots = @(
        "themePrimary", "themeLighterAlt", "themeLighter", "themeLight", "themeTertiary",
        "themeSecondary", "themeDarkAlt", "themeDark", "themeDarker",
        "neutralLighterAlt", "neutralLighter", "neutralLight", "neutralQuaternaryAlt",
        "neutralQuaternary", "neutralTertiaryAlt", "neutralTertiary", "neutralSecondary",
        "neutralPrimaryAlt", "neutralPrimary", "neutralDark", "black", "white"
    )

    foreach ($k in $officialSlots) {
        $hex = ""
        if ($script:SlotControls.Contains($k) -and -not [string]::IsNullOrWhiteSpace($script:SlotControls[$k].HexTextBox.Text)) {
            $hex = $script:SlotControls[$k].HexTextBox.Text.Trim()
        } elseif ($script:CurrentPalette.Contains($k)) {
            $hex = $script:CurrentPalette[$k]
        } else {
            $hex = "#000000"
        }
        if ($hex -notmatch '^#?[0-9a-fA-F]{6}$') { $hex = "#000000" }
        if (-not $hex.StartsWith("#")) { $hex = "#$hex" }
        $paddedKey = ('"' + $k + '"').PadRight(24)
        $sb.AppendLine("    $paddedKey = `"$($hex.ToLower())`"") | Out-Null
    }
    $sb.AppendLine("}") | Out-Null

    $sb.AppendLine("") | Out-Null

    $sb.AppendLine("# 1. Connect to SharePoint Online Admin Center") | Out-Null

    

    if (-not [string]::IsNullOrWhiteSpace($clientId) -and -not [string]::IsNullOrWhiteSpace($clientSecret)) {

        $sb.AppendLine("# App-Only Service Principal Authentication") | Out-Null

        $sb.AppendLine("if (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue) {") | Out-Null

        $sb.AppendLine("    Write-Host `"Connecting using Entra ID App Credentials to `$AdminCenterUrl...`" -ForegroundColor Cyan") | Out-Null

        $sb.AppendLine("    Connect-PnPOnline -Url `$AdminCenterUrl -ClientId `$ClientId -ClientSecret `$ClientSecret -Tenant `$TenantId") | Out-Null

        $sb.AppendLine("} else {") | Out-Null

        $sb.AppendLine("    Write-Host `"Connecting to `$AdminCenterUrl via Connect-SPOService...`" -ForegroundColor Cyan") | Out-Null

        $sb.AppendLine("    Connect-SPOService -Url `$AdminCenterUrl") | Out-Null

        $sb.AppendLine("}") | Out-Null

    } else {

        $sb.AppendLine("Write-Host `"Connecting interactively to `$AdminCenterUrl...`" -ForegroundColor Cyan") | Out-Null

        $sb.AppendLine("Connect-SPOService -Url `$AdminCenterUrl") | Out-Null

    }



    $sb.AppendLine("") | Out-Null

    $sb.AppendLine("# 2. Add or Update Theme in Tenant Catalog") | Out-Null

    $sb.AppendLine("Write-Host `"Registering theme '`$ThemeName' in tenant catalog...`" -ForegroundColor Cyan") | Out-Null

    $sb.AppendLine("try {") | Out-Null

    $sb.AppendLine("    # Try registering fresh theme") | Out-Null

    $sb.AppendLine("    Add-SPOTheme -Identity `$ThemeName -Palette `$ThemePalette -IsInverted $inv -ErrorAction Stop") | Out-Null

    $sb.AppendLine("    Write-Host `"Theme '`$ThemeName' successfully registered in tenant.`" -ForegroundColor Green") | Out-Null

    $sb.AppendLine("} catch {") | Out-Null

    $sb.AppendLine("    # If already exists, overwrite") | Out-Null

    $sb.AppendLine("    Write-Host `"Theme '`$ThemeName' already exists in tenant. Updating with -Overwrite...`" -ForegroundColor Yellow") | Out-Null

    $sb.AppendLine("    Add-SPOTheme -Identity `$ThemeName -Palette `$ThemePalette -IsInverted $inv -Overwrite") | Out-Null

    $sb.AppendLine("    Write-Host `"Theme '`$ThemeName' successfully updated.`" -ForegroundColor Green") | Out-Null

    $sb.AppendLine("}") | Out-Null

    $sb.AppendLine("") | Out-Null

    $sb.AppendLine("# 3. Apply Theme to target site") | Out-Null

    $sb.AppendLine("if (-not [string]::IsNullOrWhiteSpace(`$TargetSiteUrl)) {") | Out-Null

    $sb.AppendLine("    Write-Host `"Applying '`$ThemeName' to `$TargetSiteUrl...`" -ForegroundColor Cyan") | Out-Null

    $sb.AppendLine("    `$applied = `$false") | Out-Null

    $sb.AppendLine("    for (`$attempt = 1; `$attempt -le 4; `$attempt++) {") | Out-Null

    $sb.AppendLine("        try {") | Out-Null

    $sb.AppendLine("            Write-Host `"Applying '`$ThemeName' to `$TargetSiteUrl (Attempt `$attempt of 4)...`" -ForegroundColor Cyan") | Out-Null

    $sb.AppendLine("            Set-SPOWebTheme -Theme `$ThemeName -Web `$TargetSiteUrl -ErrorAction Stop") | Out-Null

    $sb.AppendLine("            `$applied = `$true") | Out-Null

    $sb.AppendLine("            Write-Host `"Success! Theme '`$ThemeName' is now active on `$TargetSiteUrl.`" -ForegroundColor Green") | Out-Null

    $sb.AppendLine("            Write-Host `"Tip: Hard refresh in your browser (Ctrl + F5) to load the new theme tokens.`" -ForegroundColor Yellow") | Out-Null

    $sb.AppendLine("            break") | Out-Null

    $sb.AppendLine("        } catch {") | Out-Null

    $sb.AppendLine("            if (`$_.Exception.Message -like `"*ThemeNotFoundWhenSetThemeToWeb*`" -and `$attempt -lt 4) {") | Out-Null

    $sb.AppendLine("                Write-Host `"Tenant catalog replication pending. Retrying in 3 seconds...`" -ForegroundColor Yellow") | Out-Null

    $sb.AppendLine("                Start-Sleep -Seconds 3") | Out-Null

    $sb.AppendLine("            } else {") | Out-Null

    $sb.AppendLine("                Write-Host `"Failed to apply theme to `$TargetSiteUrl : `$_`" -ForegroundColor Red") | Out-Null

    $sb.AppendLine("                Write-Host `"Tip: Verify the site URL in your browser. Common SharePoint site paths include '/sites/yourSite' or '/teams/yourSite'.`" -ForegroundColor Yellow") | Out-Null

    $sb.AppendLine("                Write-Host `"Note: The theme is registered in tenant catalog! You can also apply it via SharePoint web UI: Settings -> Change the look -> Theme.`" -ForegroundColor Cyan") | Out-Null

    $sb.AppendLine("                break") | Out-Null

    $sb.AppendLine("            }") | Out-Null

    $sb.AppendLine("        }") | Out-Null

    $sb.AppendLine("    }") | Out-Null

    $sb.AppendLine("} else {") | Out-Null

    $sb.AppendLine("    Write-Host `"Theme '`$ThemeName' is registered. You can apply it in site settings under 'Change the look'.`" -ForegroundColor Green") | Out-Null

    $sb.AppendLine("}") | Out-Null



    return $sb.ToString()

}



function Generate-JsonPalette() {

    $dict = [ordered]@{}

    foreach ($k in $script:SlotMetadata.Keys) {

        $dict[$k] = if ($script:CurrentPalette.Contains($k)) { $script:CurrentPalette[$k] } else { "#000000" }

    }

    return ($dict | ConvertTo-Json -Depth 2)

}



# Copy PowerShell Script

$copyScriptAction = {

    $scriptText = Generate-PowerShellScript

    [System.Windows.Clipboard]::SetText($scriptText)

    $txtStatusMessage.Text = "PowerShell Add-SPOTheme script copied to clipboard!"

    [System.Windows.MessageBox]::Show("PowerShell Add-SPOTheme deployment script has been copied to your clipboard!", "Copied to Clipboard", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

}



$btnTopCopyScript.Add_Click($copyScriptAction)

$btnCopyPowerShell.Add_Click($copyScriptAction)



# Save Script File Dialog

$btnSaveScript.Add_Click({

    $rawThemeName = $txtThemeName.Text.Trim()

    $themeName = ($rawThemeName -replace '[#%&*{}\\:<>\?\/+\|\"\(\)\[\]]', '') -replace '\s+', '_'

    if ([string]::IsNullOrEmpty($themeName)) { $themeName = "Custom_SPO_Theme" }

    

    $sfd = [System.Windows.Forms.SaveFileDialog]::new()

    $sfd.Title = "Save SharePoint Online Theme Deployment Script"

    $sfd.Filter = "PowerShell Script (*.ps1)|*.ps1|All Files (*.*)|*.*"

    $sfd.FileName = "Apply-$themeName.ps1"



    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

        $content = Generate-PowerShellScript

        $utf8WithBom = [System.Text.UTF8Encoding]::new($true)

        [System.IO.File]::WriteAllText($sfd.FileName, $content, $utf8WithBom)

        $txtStatusMessage.Text = "Saved script to $($sfd.FileName)."

        [System.Windows.MessageBox]::Show("Deployment script successfully saved to:`n$($sfd.FileName)", "Script Saved", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

    }

})



# Copy JSON Palette

$btnCopyJson.Add_Click({

    $json = Generate-JsonPalette

    [System.Windows.Clipboard]::SetText($json)

    $txtStatusMessage.Text = "JSON palette copied to clipboard!"

    [System.Windows.MessageBox]::Show("SharePoint Theme Palette JSON copied to clipboard!", "JSON Copied", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

})



# Direct Deploy to Tenant

$btnDeployNow.Add_Click({
    # Check for required PowerShell module
    $spoModule = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
    if ($null -eq $spoModule) {
        $warnMsg = "SharePoint Online Management Shell ('Microsoft.Online.SharePoint.PowerShell') was not found on this computer.`n`nTo run this script, install it via:`nInstall-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser`n`nDo you want to proceed anyway?"
        $choice = [System.Windows.MessageBox]::Show($warnMsg, "Module Prerequisite Check", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }

    $adminUrl = $txtAdminUrl.Text.Trim()
    $rawThemeName = $txtThemeName.Text.Trim()
    $themeName = ($rawThemeName -replace '[#%&*{}\\:<>\?\/+\|\"\(\)\[\]''`;]', '') -replace '\s+', ' '
    $siteUrl = $txtSiteUrl.Text.Trim()

    $msg = "Do you want to deploy '$themeName' to your SharePoint Online tenant?`n`nAdmin URL: $adminUrl`nSite URL: $siteUrl`n`nEnsure SharePoint Online Management Shell is installed."

    $res = [System.Windows.MessageBox]::Show($msg, "Deploy Theme to Tenant", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)



    if ($res -eq [System.Windows.MessageBoxResult]::Yes) {

        $scriptFileName = "Deploy-" + ($themeName -replace '[^\w\-]', '_') + ".ps1"

        $scriptPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $scriptFileName)

        $code = Generate-PowerShellScript

        $utf8WithBom = [System.Text.UTF8Encoding]::new($true)

        [System.IO.File]::WriteAllText($scriptPath, $code, $utf8WithBom)



        $procInfo = [System.Diagnostics.ProcessStartInfo]::new()

        $procInfo.FileName = "powershell.exe"

        $procInfo.Arguments = "-NoExit -ExecutionPolicy Bypass -File `"$scriptPath`""

        [System.Diagnostics.Process]::Start($procInfo) | Out-Null



        $txtStatusMessage.Text = "Spawned deployment session in external PowerShell console."

    }

})



# Fetch Theme From Tenant Button Click

if ($null -ne $btnFetchTenant) {

    $btnFetchTenant.Add_Click({

        $adminUrl = $txtAdminUrl.Text.Trim()

        $rawThemeName = $txtThemeName.Text.Trim()

        $themeName = ($rawThemeName -replace '[#%&*{}\:<>\?\/+\|"\(\)\[\]]', '') -replace '\s+', ' '

        $themeName = $themeName.Trim()



        if ([string]::IsNullOrWhiteSpace($adminUrl) -or [string]::IsNullOrWhiteSpace($themeName)) {

            [System.Windows.MessageBox]::Show("Please ensure Admin URL and Theme Identity are filled.", "Missing Parameters", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)

            return

        }



        try {

            $txtStatusMessage.Text = "Connecting to $adminUrl and reading '$themeName'..."

            Connect-SPOService -Url $adminUrl -ErrorAction Stop

            $themeObj = Get-SPOTheme -Identity $themeName -ErrorAction Stop

            if ($null -ne $themeObj -and $null -ne $themeObj.Palette) {

                Load-PaletteDictionary $themeObj.Palette $themeObj.IsInverted

                $txtStatusMessage.Text = "Successfully synchronized form with live tenant theme '$themeName'!"

                [System.Windows.MessageBox]::Show("Successfully fetched theme '$themeName' from tenant! All 22 slot hexcodes and preview are now aligned.", "Tenant Theme Loaded", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

            } else {

                [System.Windows.MessageBox]::Show("Theme '$themeName' exists but returned an empty palette.", "Empty Palette", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)

            }

        } catch {

            [System.Windows.MessageBox]::Show("Could not fetch theme '$themeName' from ${adminUrl}:`n`n$_", "Fetch Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)

        }

    })

}



# ---------------------------------------------------------------------------

# INITIALIZATION & EXECUTION

# ---------------------------------------------------------------------------



Build-SlotFormRows

Register-PreviewInteractivity



# Load initial preset

$defaultPreset = $script:Presets["Microsoft Modern Blue"]

Load-PaletteDictionary $defaultPreset.Slots $defaultPreset.IsInverted

Update-Inspector "themePrimary"



# Show the GUI Window

$window.ShowDialog() | Out-Null

