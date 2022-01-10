<?php 
namespace Waka\OpenTBS;

use clsTinyButStrong;

class MergePpt
{
    public static clsTinyButStrong $tbs; 
    protected $modelPath;
    private $varRef;

    public function __construct() {
        self::$tbs = new clsTinyButStrong;
        $ok = class_exists('clsOpenTBS', true);
        self::$tbs->Plugin(TBS_INSTALL,  OPENTBS_PLUGIN);
        self::$tbs->ResetVarRef(false); // VarRef is now a new empty array
    }

    public function loadTemplate($src) {
        self::$tbs->LoadTemplate($src, OPENTBS_ALREADY_UTF8); // Also merge some [onload] automatic fields (depends of the type of document).
        return $this;
    }

    public function degubTemplate() {
        //return self::$tbs->PlugIn(OPENTBS_DEBUG_INFO ,[$Exit]);
        //return self::$tbs->PlugIn(OPENTBS_DEBUG_XML_SHOW);
        return self::$tbs->PlugIn(OPENTBS_DEBUG_XML_CURRENT, true);

    }


    public function mergeField($slideNum, array $data = [], $prefix = null) {
        self::$tbs->Plugin(OPENTBS_SELECT_SLIDE, $slideNum);
        if(!$prefix) {
            $prefix = 'ds';
        }
        self::$tbs->MergeField($prefix, $data);
        return $this;
    }

    public function MergeBlock($slideNum, $key, array $data = []) {
        self::$tbs->Plugin(OPENTBS_SELECT_SLIDE, $slideNum);
        self::$tbs->MergeBlock($key, $data);
        return $this;
    }

    public function changeSlide($slideNum) {
        self::$tbs->Plugin(OPENTBS_SELECT_SLIDE, $slideNum);
    }

    
    public function changePicture($slideNum,$pictureRef, $pictureSrc, $_options = []) {
        $baseOptions = ['unique' => 1, 'adjust' => 'inside'];
        $options = array_merge($baseOptions, $_options);
        self::$tbs->Plugin(OPENTBS_SELECT_SLIDE, $slideNum);
        self::$tbs->Plugin(OPENTBS_CHANGE_PICTURE, $pictureRef, $pictureSrc, $options);
    }

    public function changeChart($ChartRef, $series, $overrideLegend = false ) {
        foreach($series as $key=>$serie) {
            $label = $serie['label'] ?? false;
            $datas = $serie['datas'] ?? [];
            $delete = $serie['delete'] ?? false;
            if($delete) {
                self::$tbs->PlugIn(OPENTBS_CHART_DELETE_CATEGORY, $ChartRef, $key, $NoErr = false);
            } elseif($datas) {
                 self::$tbs->PlugIn(OPENTBS_CHART, $ChartRef, $key, $datas, $label);
            }
        }
    }

    public function deleteSlide($NumOrNames) {
        self::$tbs->PlugIn(OPENTBS_DELETE_SLIDES, $NumOrNames);
    }


    public function savePpt($path) {
        // //Chargement des variables dans l'objet VarRef. Permet de ne pas utiliser les variables globales. 
        // self::$tbs->VarRef = $this->varRef;
        // // Define the name of the output file
        $output_file_name = $path;
        // Output the result as a downloadable file (only streaming, no data saved in the server)
        self::$tbs->Show(OPENTBS_FILE, $output_file_name); // Also merges all [onshow] automatic fields.
        // Be sure that no more output is done, otherwise the download file is corrupted with extra data.
        exit("File [$output_file_name] has been created.");

    }

    public function downloadPpt($name) {
        // //Chargement des variables dans l'objet VarRef. Permet de ne pas utiliser les variables globales. 
        // self::$tbs->VarRef = $this->varRef;
        // // Define the name of the output file
        $output_file_name = $name;
        // Output the result as a downloadable file (only streaming, no data saved in the server)
        self::$tbs->Show(OPENTBS_DOWNLOAD, $output_file_name); // Also merges all [onshow] automatic fields.
        // Be sure that no more output is done, otherwise the download file is corrupted with extra data.
        exit();
    }
}