
#include <iostream>
using namespace std;


    static char map[10][10] =  {    {'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X' }, 
                                    {'X', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', 'X' } ,
                                    {'X', ' ', 'X', 'X', ' ', 'X', ' ', 'X', ' ', 'X' } ,
                                    {'X', ' ', 'X', 'X', ' ', 'X', ' ', 'X', ' ', 'X' } ,
                                    {'X', ' ', 'X', 'X', ' ', 'X', ' ', 'X', ' ', 'X' } ,
                                    {'X', ' ', ' ', ' ', ' ', ' ', ' ', 'X', ' ', 'E' } ,
                                    {'X', ' ', 'X', 'X', ' ', 'X', ' ', 'X', ' ', 'X' } ,
                                    {'X', ' ', 'X', 'X', ' ', 'X', ' ', 'X', ' ', 'X' } ,
                                    {'X', ' ', 'X', 'X', ' ', ' ', ' ', ' ', ' ', 'X' } ,
                                    {'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X' } 
                            };


int main(void){ 
    int heroPosition[2] = {1,1};
    int dragonPosition[2] = {3,1};
    int keyPosition[2] = {8,1};
    int exitPosition[2] = {5,9};

    bool game = true;
    bool key = false;
    char walk;

    bool winLose = false;



    while(game){

                map[heroPosition[0]][heroPosition[1]] = 'H';
                map[dragonPosition[0]][dragonPosition[1]] = 'D';
                if(!key){
                    
                    
                
                    
                
                    map[keyPosition[0]][keyPosition[1]] = 'K';

                }
            for(int i = 0; i < 10; i++){
                for(int j = 0; j < 10; j++){
                        cout << map[i][j];
                        cout << " ";

                }
                cout << "\n";
            }

            cout << "Please walk\n";    
            cin >> walk;

            

            switch(walk){
                case('w'):
                    if(map[heroPosition[0]-1][heroPosition[1]] == 'X' || (map[heroPosition[0]-1][heroPosition[1]] == 'E' && key == false) ){
                        cout<<"Can't go there!\n";
                    }else{
                        map[heroPosition[0]][heroPosition[1]] = ' ';
                        heroPosition[0] -= 1;
                    }
                    break;
                case('s'):
                    if(map[heroPosition[0]+1][heroPosition[1]] == 'X' || (map[heroPosition[0]+1][heroPosition[1]] == 'E' && key == false) ){
                        cout<<"Can't go there!\n";
                    }else{
                        map[heroPosition[0]][heroPosition[1]] = ' ';
                        heroPosition[0] += 1;
                    }

                    break;
                case('a'):
                    if(map[heroPosition[0]][heroPosition[1]-1] == 'X' || (map[heroPosition[0]][heroPosition[1]-1] == 'E' && key == false) ){
                        cout<<"Can't go there!\n";
                    }else{
                        map[heroPosition[0]][heroPosition[1]] = ' ';
                        heroPosition[1] -= 1;
                    }
                    break;
                case('d'):
                    if(map[heroPosition[0]][heroPosition[1]+1] == 'X' || (map[heroPosition[0]][heroPosition[1]+1] == 'E' && key == false) ){
                        cout<<"Can't go there!\n";
                    }else{
                        map[heroPosition[0]][heroPosition[1]] = ' ';
                        heroPosition[1] += 1;
                    }
                    break;
                default:
                    cout << "Wrong key!\n";
                    break;


            }

            if(heroPosition[0] == keyPosition[0] && heroPosition[1] == keyPosition[1]){
                key = true;
                 map[keyPosition[0]][keyPosition[1]] = 'H';
            }

            if(key == true && heroPosition[0] == exitPosition[0] && heroPosition[1] == exitPosition[1]){
                game = false;
                winLose = true;
            }   
            
            for(int i = -1; i < 1; i++){
                for(int j = -1; j < 1; j++){
                    cout << (dragonPosition[0]+i);
                    cout << " ";
                    cout << (dragonPosition[1]+i);
                    cout << "\n";
                    if(dragonPosition[0]+i == heroPosition[0] && dragonPosition[1]+j == heroPosition[1]){
                        
                        game = false;
                        winLose = false;
                    }
                }
            }
            

    }

    if(winLose){
        cout << "\n\nVICTORY!\n\n";
    }else{
        cout << "\n\nGAME OVER!\n\n";
    }   




} 
